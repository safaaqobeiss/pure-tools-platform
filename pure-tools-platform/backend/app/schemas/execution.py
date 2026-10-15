from datetime import datetime

from pydantic import BaseModel


class ExecutionResponse(BaseModel):
    id: str
    user_id: int
    tool_id: str
    status: str
    parameters: str | None
    error_message: str | None
    created_at: datetime

    model_config = {"from_attributes": True}


class ExecutionListResponse(BaseModel):
    id: str
    tool_id: str
    tool_name: str | None = None
    script_path: str | None = None
    status: str
    created_at: datetime
    error_message: str | None
    user_name: str | None = None
    user_email: str | None = None

    model_config = {"from_attributes": True}
