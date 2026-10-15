from fastapi import APIRouter, Depends, File, HTTPException, UploadFile, status
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.middleware.auth_middleware import get_current_user
from app.models.user import User
from app.schemas.user import Token, UserLogin, UserRegister, UserResponse, UserUpdate
from app.services.auth_service import (
    authenticate_user,
    create_access_token,
    hash_password,
    register_user,
    verify_password,
)

router = APIRouter(prefix="/api/auth", tags=["Authentication"])


@router.post("/register", response_model=UserResponse, status_code=status.HTTP_201_CREATED)
def register(user_data: UserRegister, db: Session = Depends(get_db)):
    try:
        user = register_user(db, user_data)
        return user
    except ValueError as e:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(e))


@router.post("/login", response_model=Token)
def login(user_data: UserLogin, db: Session = Depends(get_db)):
    user = authenticate_user(db, user_data.email, user_data.password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid email or password",
        )
    access_token = create_access_token(user.id)
    return Token(access_token=access_token)


@router.get("/me", response_model=UserResponse)
def get_me(current_user: User = Depends(get_current_user)):
    return _user_response(current_user)


@router.put("/profile", response_model=UserResponse)
def update_profile(
    data: UserUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Update user's name and/or password."""
    if data.full_name is not None:
        current_user.full_name = data.full_name

    if data.new_password:
        if not data.current_password:
            raise HTTPException(status_code=400, detail="Mot de passe actuel requis")
        if not verify_password(data.current_password, current_user.password):
            raise HTTPException(status_code=400, detail="Mot de passe actuel incorrect")
        current_user.password = hash_password(data.new_password)

    db.commit()
    db.refresh(current_user)
    return _user_response(current_user)


@router.post("/avatar", response_model=UserResponse)
async def upload_avatar(
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Upload a user avatar image."""
    from pathlib import Path

    if file.content_type not in ("image/png", "image/jpeg", "image/webp", "image/gif"):
        raise HTTPException(status_code=400, detail="Format non supporté (png, jpg, webp, gif)")

    avatars_dir = Path(__file__).resolve().parent.parent.parent / "storage" / "avatars"
    avatars_dir.mkdir(parents=True, exist_ok=True)

    ext = file.filename.rsplit(".", 1)[-1] if "." in file.filename else "png"
    avatar_filename = f"avatar_{current_user.id}.{ext}"
    avatar_path = avatars_dir / avatar_filename

    # Remove old avatar files with different extensions
    for old_ext in ("png", "jpg", "jpeg", "webp", "gif"):
        old_path = avatars_dir / f"avatar_{current_user.id}.{old_ext}"
        if old_path.is_file() and old_path != avatar_path:
            old_path.unlink()

    content = await file.read()
    if len(content) > 2 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="Fichier trop volumineux (max 2 Mo)")
    with open(avatar_path, "wb") as f:
        f.write(content)

    current_user.avatar_path = str(avatar_path)
    db.commit()
    db.refresh(current_user)
    return _user_response(current_user)


@router.get("/avatar/{user_id}")
def get_avatar(user_id: int):
    """Serve the avatar image for a user."""
    from pathlib import Path

    avatars_dir = Path(__file__).resolve().parent.parent.parent / "storage" / "avatars"
    for ext in ("png", "jpg", "jpeg", "webp", "gif"):
        path = avatars_dir / f"avatar_{user_id}.{ext}"
        if path.is_file():
            return FileResponse(
                path,
                headers={"Cache-Control": "no-cache, no-store, must-revalidate"},
            )
    raise HTTPException(status_code=404, detail="Avatar not found")


def _user_response(user: User) -> dict:
    return {
        "id": user.id,
        "email": user.email,
        "full_name": user.full_name,
        "avatar_url": f"/api/auth/avatar/{user.id}" if user.avatar_path else None,
        "role": user.role,
        "is_active": user.is_active,
    }
