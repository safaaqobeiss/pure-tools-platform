"""Dashboard stats endpoint."""

from datetime import datetime, timedelta, timezone
from io import BytesIO

from fastapi import APIRouter, Depends
from fastapi.responses import StreamingResponse
from fpdf import FPDF
from sqlalchemy import func, text
from sqlalchemy.orm import Session

from app.database import get_db
from app.middleware.auth_middleware import get_current_user
from app.models.tool import Tool
from app.models.tool_execution import ToolExecution
from app.models.user import User

router = APIRouter(prefix="/api/dashboard", tags=["Dashboard"])


@router.get("/stats")
def get_stats(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Return dashboard statistics."""
    # Total executions
    total = db.query(func.count(ToolExecution.id)).scalar() or 0
    completed = (
        db.query(func.count(ToolExecution.id))
        .filter(ToolExecution.status == "completed")
        .scalar()
        or 0
    )
    failed = (
        db.query(func.count(ToolExecution.id))
        .filter(ToolExecution.status == "failed")
        .scalar()
        or 0
    )
    success_rate = round((completed / total * 100), 1) if total > 0 else 0

    # Total users
    total_users = db.query(func.count(User.id)).scalar() or 0

    # Executions per tool
    tool_stats = (
        db.query(Tool.name, func.count(ToolExecution.id))
        .join(ToolExecution, Tool.id == ToolExecution.tool_id)
        .group_by(Tool.name)
        .all()
    )
    per_tool = [{"tool": name, "count": count} for name, count in tool_stats]

    # Most used tool
    most_used = max(per_tool, key=lambda x: x["count"])["tool"] if per_tool else None

    # Executions per day (last 14 days)
    cutoff = datetime.now(timezone.utc) - timedelta(days=14)
    from sqlalchemy import text

    raw = db.execute(
        text(
            "SELECT CAST(created_at AS DATE) as d, COUNT(*) as c "
            "FROM tool_executions "
            "WHERE created_at >= :cutoff "
            "GROUP BY CAST(created_at AS DATE) "
            "ORDER BY d"
        ),
        {"cutoff": cutoff},
    )
    per_day = [{"date": str(row[0]), "count": row[1]} for row in raw]

    # Executions per status
    per_status = [
        {"status": "completed", "count": completed},
        {"status": "failed", "count": failed},
        {"status": "other", "count": total - completed - failed},
    ]

    # My executions
    my_total = (
        db.query(func.count(ToolExecution.id))
        .filter(ToolExecution.user_id == current_user.id)
        .scalar()
        or 0
    )

    return {
        "total_executions": total,
        "completed": completed,
        "failed": failed,
        "success_rate": success_rate,
        "total_users": total_users,
        "most_used_tool": most_used,
        "my_executions": my_total,
        "per_tool": per_tool,
        "per_day": per_day,
        "per_status": per_status,
    }


@router.get("/export")
def export_stats_pdf(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Export dashboard statistics as a PDF report."""
    stats = get_stats(db=db, current_user=current_user)
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # --- Decorative header bar ---
    pdf.set_fill_color(108, 44, 145)  # #6c2c91
    pdf.rect(0, 0, 210, 45, "F")
    # Accent stripe
    pdf.set_fill_color(231, 76, 138)  # #e74c8a
    pdf.rect(0, 45, 210, 3, "F")

    # --- Title on header ---
    pdf.set_y(10)
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 12, "Pure Tools Platform", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(220, 200, 240)
    pdf.cell(0, 7, f"Rapport du tableau de bord - {now}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, f"Genere par : {current_user.full_name or current_user.email}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(12)

    # --- Section: KPI ---
    _section_title(pdf, "Indicateurs cles")

    kpis = [
        ("Total executions", str(stats["total_executions"])),
        ("Executions reussies", str(stats["completed"])),
        ("Executions echouees", str(stats["failed"])),
        ("Taux de succes", f"{stats['success_rate']}%"),
        ("Nombre d'utilisateurs", str(stats["total_users"])),
        ("Outil le plus utilise", stats["most_used_tool"] or "-"),
        ("Mes executions", str(stats["my_executions"])),
    ]

    col_w = [100, 80]
    # Header
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(108, 44, 145)  # #6c2c91
    pdf.set_text_color(255, 255, 255)
    pdf.cell(col_w[0], 8, "Indicateur", border=1, fill=True)
    pdf.cell(col_w[1], 8, "Valeur", border=1, fill=True, new_x="LMARGIN", new_y="NEXT")
    # Rows
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(50, 50, 50)
    for i, (label, value) in enumerate(kpis):
        bg = i % 2 == 0
        if bg:
            pdf.set_fill_color(245, 235, 250)  # light purple
        pdf.cell(col_w[0], 7, label, border=1, fill=bg)
        pdf.cell(col_w[1], 7, value, border=1, fill=bg, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)

    # --- Section: Per tool ---
    if stats["per_tool"]:
        _section_title(pdf, "Repartition par outil")
        col_w2 = [120, 60]
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_fill_color(155, 58, 177)  # #9b3ab1
        pdf.set_text_color(255, 255, 255)
        pdf.cell(col_w2[0], 8, "Outil", border=1, fill=True)
        pdf.cell(col_w2[1], 8, "Executions", border=1, fill=True, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(50, 50, 50)
        for i, item in enumerate(stats["per_tool"]):
            bg = i % 2 == 0
            if bg:
                pdf.set_fill_color(248, 235, 255)  # lighter purple
            pdf.cell(col_w2[0], 7, item["tool"], border=1, fill=bg)
            pdf.cell(col_w2[1], 7, str(item["count"]), border=1, fill=bg, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(8)

    # --- Section: Per day ---
    if stats["per_day"]:
        _section_title(pdf, "Executions par jour (14 derniers jours)")
        col_w3 = [100, 80]
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_fill_color(231, 76, 138)  # #e74c8a
        pdf.set_text_color(255, 255, 255)
        pdf.cell(col_w3[0], 8, "Date", border=1, fill=True)
        pdf.cell(col_w3[1], 8, "Nombre", border=1, fill=True, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(50, 50, 50)
        for i, item in enumerate(stats["per_day"]):
            bg = i % 2 == 0
            if bg:
                pdf.set_fill_color(253, 235, 245)  # light pink
            pdf.cell(col_w3[0], 7, item["date"], border=1, fill=bg)
            pdf.cell(col_w3[1], 7, str(item["count"]), border=1, fill=bg, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(8)

    # --- Footer ---
    pdf.set_y(-25)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(150, 150, 150)
    pdf.cell(0, 10, f"Pure Tools Platform - Rapport genere le {now}", align="C")

    buf = BytesIO()
    pdf.output(buf)
    buf.seek(0)

    filename = f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


def _section_title(pdf: FPDF, title: str):
    """Render a styled section title."""
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(108, 44, 145)  # #6c2c91
    pdf.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(155, 58, 177)  # #9b3ab1
    pdf.set_line_width(0.6)
    pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 180, pdf.get_y())
    pdf.set_line_width(0.2)
    pdf.ln(4)
