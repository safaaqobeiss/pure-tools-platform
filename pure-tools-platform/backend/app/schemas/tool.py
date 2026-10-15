from datetime import datetime

from pydantic import BaseModel


class ToolCreate(BaseModel):
    name: str
    description: str | None = None
    icon: str | None = None
    category: str | None = None
    script_path: str


class ToolUpdate(BaseModel):
    name: str | None = None
    description: str | None = None
    icon: str | None = None
    category: str | None = None
    is_active: bool | None = None


class ToolResponse(BaseModel):
    id: str
    name: str
    description: str | None
    icon: str | None
    category: str | None
    script_path: str
    is_active: bool
    created_at: datetime

    model_config = {"from_attributes": True}
