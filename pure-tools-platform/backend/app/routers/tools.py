from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from app.database import get_db
from app.middleware.auth_middleware import get_current_user
from app.models.tool import Tool
from app.models.user import User
from app.schemas.tool import ToolCreate, ToolResponse, ToolUpdate

router = APIRouter(prefix="/api/tools", tags=["Tools"])


@router.get("/", response_model=list[ToolResponse])
def list_tools(
    db: Session = Depends(get_db),
    _current_user: User = Depends(get_current_user),
):
    tools = db.query(Tool).filter(Tool.is_active == True).all()
    return tools


@router.get("/{tool_id}", response_model=ToolResponse)
def get_tool(
    tool_id: str,
    db: Session = Depends(get_db),
    _current_user: User = Depends(get_current_user),
):
    tool = db.query(Tool).filter(Tool.id == tool_id).first()
    if not tool:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Tool not found")
    return tool


@router.post("/", response_model=ToolResponse, status_code=status.HTTP_201_CREATED)
def create_tool(
    tool_data: ToolCreate,
    db: Session = Depends(get_db),
    _current_user: User = Depends(get_current_user),
):
    tool = Tool(**tool_data.model_dump())
    db.add(tool)
    db.commit()
    db.refresh(tool)
    return tool


@router.put("/{tool_id}", response_model=ToolResponse)
def update_tool(
    tool_id: str,
    tool_data: ToolUpdate,
    db: Session = Depends(get_db),
    _current_user: User = Depends(get_current_user),
):
    tool = db.query(Tool).filter(Tool.id == tool_id).first()
    if not tool:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Tool not found")
    for key, value in tool_data.model_dump(exclude_unset=True).items():
        setattr(tool, key, value)
    db.commit()
    db.refresh(tool)
    return tool


@router.delete("/{tool_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_tool(
    tool_id: str,
    db: Session = Depends(get_db),
    _current_user: User = Depends(get_current_user),
):
    tool = db.query(Tool).filter(Tool.id == tool_id).first()
    if not tool:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Tool not found")
    db.delete(tool)
    db.commit()
