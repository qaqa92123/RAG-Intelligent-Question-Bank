from threading import Lock

from langchain_milvus import Milvus

from config import RagConfig
from text_utils.text_embeddings import RagTextEmbeddings

# 加载配置
config = RagConfig()

# 配置索引参数和搜索参数
index_params = {
    "index_type": "IVF_FLAT",
    "metric_type": "L2",
    "params": {"nlist": 100}
}

search_params = {
    "metric_type": "L2",
    "params": {"nprobe": 10}
}

_embedding_instance = None
_vector_store_instance = None
_embedding_lock = Lock()
_vector_store_lock = Lock()


def get_embeddings():
    """获取进程级单例 embedding 实例，避免重复加载 SentenceTransformer。"""
    global _embedding_instance
    if _embedding_instance is None:
        with _embedding_lock:
            if _embedding_instance is None:
                _embedding_instance = RagTextEmbeddings(
                    embed_model_path=config.text_embeddings_model_path,
                    batch_size=32,
                    device=config.device,
                )
    return _embedding_instance


def get_vector_store():
    """
    获取进程级单例 Milvus 向量存储实例。
    首次调用时初始化，后续请求直接复用。
    """
    global _vector_store_instance
    if _vector_store_instance is None:
        with _vector_store_lock:
            if _vector_store_instance is None:
                _vector_store_instance = Milvus(
                    embedding_function=get_embeddings(),
                    collection_name=config.milvus_collection_name,
                    consistency_level="Bounded",
                    connection_args={"host": config.milvus_host, "port": config.milvus_port},
                    index_params=index_params,
                    search_params=search_params,
                    # auto_id=True 会自动生成ID，drop_old=False 不会删除已存在的 collection
                    drop_old=False
                )
    return _vector_store_instance

# 注意：不要在模块导入时初始化 heavy 对象（如模型加载或网络下载），
# 请通过调用 `get_vector_store()` / `get_embeddings()` 在需要时懒加载单例实例。

