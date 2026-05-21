from pydantic import BaseModel
class ChatRequest(BaseModel):
    question: str


class ClearRequest(BaseModel):
    collection_name: str | None = None
    host: str | None = None
    port: int | None = None

