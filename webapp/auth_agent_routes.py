from typing import Any, Callable, Dict, Optional
from fastapi import Cookie, Form, Request, Response
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.templating import Jinja2Templates


def register_auth_and_agent_routes(
    web_app,
    templates: Jinja2Templates,
    auth_db_module: Any,
    *,
    build_agent_tool_catalog: Callable[[], list[dict[str, Any]]],
    safe_str: Callable[[Any], str],
    generate_agent_paper_task: Callable[..., Dict[str, Any]],
    load_agent_task_record: Callable[[str], Optional[Dict[str, Any]]],
    export_agent_paper_task: Callable[[Dict[str, Any], Optional[Dict[str, Any]]], Dict[str, Any]],
    build_agent_paper_request_from_text: Callable[[str, Dict[str, Any]], Dict[str, Any]],
    build_agent_trace_entry: Callable[[str, str, str], Dict[str, str]],
):
    def get_current_user(session_id: Optional[str] = Cookie(None)) -> Optional[Dict]:
        if not session_id:
            return None
        return auth_db_module.verify_session(session_id)

    @web_app.get('/login', response_class=HTMLResponse)
    async def login_page(request: Request, error: str = None, success: str = None):
        return templates.TemplateResponse('login.html', {
            'request': request,
            'error': error,
            'success': success,
        })

    @web_app.post('/login')
    async def login(
        request: Request,
        response: Response,
        email: str = Form(...),
        password: str = Form(...),
    ):
        result = auth_db_module.verify_user(email, password)

        if result['success']:
            session_id = auth_db_module.create_session(result['user_id'])
            response = RedirectResponse(url='/', status_code=303)
            response.set_cookie(
                key='session_id',
                value=session_id,
                max_age=7 * 24 * 60 * 60,
                httponly=True,
                samesite='lax',
            )
            return response

        return templates.TemplateResponse('login.html', {
            'request': request,
            'error': result['message'],
        })

    @web_app.get('/register', response_class=HTMLResponse)
    async def register_page(request: Request, error: str = None):
        return templates.TemplateResponse('register.html', {
            'request': request,
            'error': error,
        })

    @web_app.post('/register')
    async def register(
        request: Request,
        username: str = Form(...),
        email: str = Form(...),
        password: str = Form(...),
    ):
        result = auth_db_module.create_user(email, password, username)

        if result['success']:
            return RedirectResponse(
                url='/login?success=注册成功，请登录',
                status_code=303,
            )

        return templates.TemplateResponse('register.html', {
            'request': request,
            'error': result['message'],
        })

    @web_app.get('/logout')
    async def logout(response: Response, session_id: Optional[str] = Cookie(None)):
        if session_id:
            auth_db_module.delete_session(session_id)

        response = RedirectResponse(url='/login?success=已成功退出', status_code=303)
        response.delete_cookie('session_id')
        return response

    @web_app.get('/agent-workbench', response_class=HTMLResponse)
    async def agent_workbench_page(request: Request, session_id: Optional[str] = Cookie(None)):
        user = get_current_user(session_id)
        if not user:
            return RedirectResponse(url='/login', status_code=303)
        return templates.TemplateResponse('agent_workspace.html', {
            'request': request,
            'user': user,
        })

    @web_app.get('/api/agent/tools', response_class=JSONResponse)
    async def agent_tools(session_id: Optional[str] = Cookie(None)):
        user = get_current_user(session_id)
        if not user:
            return JSONResponse({'success': False, 'message': '请先登录'}, status_code=401)
        return JSONResponse({
            'success': True,
            'tools': build_agent_tool_catalog(),
        })

    @web_app.post('/api/agent/paper/generate', response_class=JSONResponse)
    async def agent_generate_paper(request: Request, session_id: Optional[str] = Cookie(None)):
        user = get_current_user(session_id)
        if not user:
            return JSONResponse({'success': False, 'message': '请先登录'}, status_code=401)

        try:
            body = await request.json()
        except Exception:
            body = {}

        result = generate_agent_paper_task(
            body,
            request_text=safe_str(body.get('request_text') or body.get('message')),
        )
        status_code = 200 if result.get('success') else 400
        return JSONResponse(result, status_code=status_code)

    @web_app.post('/api/agent/execute', response_class=JSONResponse)
    async def agent_execute_tool(request: Request, session_id: Optional[str] = Cookie(None)):
        user = get_current_user(session_id)
        if not user:
            return JSONResponse({'success': False, 'message': '请先登录'}, status_code=401)

        try:
            body = await request.json()
        except Exception:
            body = {}

        tool_name = safe_str(body.get('tool'))
        arguments = body.get('arguments', {}) or {}
        if not isinstance(arguments, dict):
            return JSONResponse({'success': False, 'message': 'arguments 必须是对象'}, status_code=400)

        if tool_name == 'paper.generate':
            result = generate_agent_paper_task(arguments, request_text=safe_str(body.get('request_text')))
            return JSONResponse(result, status_code=200 if result.get('success') else 400)

        if tool_name == 'paper.export':
            task_id = safe_str(arguments.get('task_id'))
            task_record = load_agent_task_record(task_id)
            if not task_record:
                return JSONResponse({'success': False, 'message': '任务不存在'}, status_code=404)
            result = export_agent_paper_task(task_record, arguments)
            return JSONResponse(result, status_code=200 if result.get('success') else 400)

        catalog_item = next((item for item in build_agent_tool_catalog() if item.get('name') == tool_name), None)
        if catalog_item:
            return JSONResponse({
                'success': False,
                'message': '该工具已封装为标准 API，请直接调用对应的 invoke.path。',
                'tool': catalog_item,
            }, status_code=400)

        return JSONResponse({'success': False, 'message': '未知工具名称'}, status_code=404)

    @web_app.post('/api/agent/chat', response_class=JSONResponse)
    async def agent_chat(request: Request, session_id: Optional[str] = Cookie(None)):
        user = get_current_user(session_id)
        if not user:
            return JSONResponse({'success': False, 'message': '请先登录'}, status_code=401)

        try:
            body = await request.json()
        except Exception:
            body = {}

        message = safe_str(body.get('message'))
        if not message:
            return JSONResponse({'success': False, 'message': 'message 不能为空'}, status_code=400)

        overrides = body.get('overrides', {}) or {}
        if not isinstance(overrides, dict):
            overrides = {}

        plan = build_agent_paper_request_from_text(message, overrides)
        result = generate_agent_paper_task(plan, request_text=message)
        if not result.get('success'):
            return JSONResponse({
                'success': False,
                'tool': 'paper.generate',
                'plan': plan,
                'message': result.get('message') or '智能体执行失败',
                'trace': result.get('trace', []),
                'timing': result.get('timing'),
                'response_time_ms': result.get('response_time_ms'),
                'response_time_seconds': result.get('response_time_seconds'),
            }, status_code=400)

        trace = [build_agent_trace_entry('plan', '自然语言需求已转换为 paper.generate 工具参数')]
        trace.extend(result.get('trace', []))
        return JSONResponse({
            'success': True,
            'tool': 'paper.generate',
            'plan': plan,
            'trace': trace,
            'task_id': result.get('task_id'),
            'preview': result.get('preview'),
            'export': result.get('export'),
            'normalized_request': result.get('normalized_request'),
            'timing': result.get('timing'),
            'response_time_ms': result.get('response_time_ms'),
            'response_time_seconds': result.get('response_time_seconds'),
        })

    @web_app.get('/api/agent/tasks/{task_id}', response_class=JSONResponse)
    async def agent_task_detail(task_id: str, session_id: Optional[str] = Cookie(None)):
        user = get_current_user(session_id)
        if not user:
            return JSONResponse({'success': False, 'message': '请先登录'}, status_code=401)

        task_record = load_agent_task_record(task_id)
        if not task_record:
            return JSONResponse({'success': False, 'message': '任务不存在'}, status_code=404)

        return JSONResponse({'success': True, 'task': task_record})

    @web_app.post('/api/agent/tasks/{task_id}/export', response_class=JSONResponse)
    async def agent_task_export(task_id: str, request: Request, session_id: Optional[str] = Cookie(None)):
        user = get_current_user(session_id)
        if not user:
            return JSONResponse({'success': False, 'message': '请先登录'}, status_code=401)

        task_record = load_agent_task_record(task_id)
        if not task_record:
            return JSONResponse({'success': False, 'message': '任务不存在'}, status_code=404)
        if safe_str(task_record.get('task_type')) != 'paper.generate':
            return JSONResponse({'success': False, 'message': '当前仅支持导出智能体生成的试卷任务'}, status_code=400)

        try:
            body = await request.json()
        except Exception:
            body = {}

        result = export_agent_paper_task(task_record, body)
        return JSONResponse(result, status_code=200 if result.get('success') else 400)

    return get_current_user