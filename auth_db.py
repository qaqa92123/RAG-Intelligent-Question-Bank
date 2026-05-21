# -*- coding: utf-8 -*-
"""
用户认证数据库模块
提供用户注册、登录、会话管理功能
"""

import sqlite3
import hashlib
import secrets
from datetime import datetime, timedelta
from typing import Optional, Dict

DB_PATH = "users.db"

def init_db():
    """初始化数据库，创建用户表和会话表"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # 创建用户表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            username TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            last_login TIMESTAMP
        )
    """)
    
    # 创建会话表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            session_id TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            expires_at TIMESTAMP NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    """)
    
    conn.commit()
    conn.close()

def hash_password(password: str) -> str:
    """对密码进行 SHA-256 哈希"""
    return hashlib.sha256(password.encode()).hexdigest()

def create_user(email: str, password: str, username: str) -> Dict[str, any]:
    """
    注册新用户
    返回: {'success': bool, 'message': str, 'user_id': int}
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # 检查邮箱是否已存在
        cursor.execute("SELECT id FROM users WHERE email = ?", (email,))
        if cursor.fetchone():
            conn.close()
            return {'success': False, 'message': '该邮箱已被注册'}
        
        # 插入新用户
        password_hash = hash_password(password)
        cursor.execute(
            "INSERT INTO users (email, password_hash, username) VALUES (?, ?, ?)",
            (email, password_hash, username)
        )
        user_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        return {'success': True, 'message': '注册成功', 'user_id': user_id}
    
    except Exception as e:
        return {'success': False, 'message': f'注册失败: {str(e)}'}

def verify_user(email: str, password: str) -> Dict[str, any]:
    """
    验证用户登录
    返回: {'success': bool, 'message': str, 'user_id': int, 'username': str}
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        password_hash = hash_password(password)
        cursor.execute(
            "SELECT id, username FROM users WHERE email = ? AND password_hash = ?",
            (email, password_hash)
        )
        result = cursor.fetchone()
        
        if result:
            user_id, username = result
            # 更新最后登录时间
            cursor.execute(
                "UPDATE users SET last_login = ? WHERE id = ?",
                (datetime.now(), user_id)
            )
            conn.commit()
            conn.close()
            return {
                'success': True,
                'message': '登录成功',
                'user_id': user_id,
                'username': username
            }
        else:
            conn.close()
            return {'success': False, 'message': '邮箱或密码错误'}
    
    except Exception as e:
        return {'success': False, 'message': f'登录失败: {str(e)}'}

def create_session(user_id: int) -> str:
    """
    为用户创建会话
    返回: session_id
    """
    session_id = secrets.token_urlsafe(32)
    expires_at = datetime.now() + timedelta(days=7)  # 7天有效期
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO sessions (session_id, user_id, expires_at) VALUES (?, ?, ?)",
        (session_id, user_id, expires_at)
    )
    conn.commit()
    conn.close()
    
    return session_id

def verify_session(session_id: str) -> Optional[Dict[str, any]]:
    """
    验证会话是否有效
    返回: {'user_id': int, 'username': str} 或 None
    """
    if not session_id:
        return None
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT s.user_id, u.username, s.expires_at
        FROM sessions s
        JOIN users u ON s.user_id = u.id
        WHERE s.session_id = ?
    """, (session_id,))
    
    result = cursor.fetchone()
    conn.close()
    
    if not result:
        return None
    
    user_id, username, expires_at = result
    expires_at = datetime.fromisoformat(expires_at)
    
    # 检查是否过期
    if expires_at < datetime.now():
        delete_session(session_id)
        return None
    
    return {'user_id': user_id, 'username': username}

def delete_session(session_id: str):
    """删除会话（退出登录）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sessions WHERE session_id = ?", (session_id,))
    conn.commit()
    conn.close()

def clean_expired_sessions():
    """清理过期的会话"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sessions WHERE expires_at < ?", (datetime.now(),))
    conn.commit()
    conn.close()

# 初始化数据库
init_db()
