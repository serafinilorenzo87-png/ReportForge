from datetime import datetime

from sqlalchemy import DateTime, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.services.database import Base


class Project(Base):
    __tablename__ = "projects"

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)

    project_name: Mapped[str] = mapped_column(String(255), nullable=False)
    target_name: Mapped[str] = mapped_column(String(255), default="")
    platform: Mapped[str] = mapped_column(String(255), default="")
    client_name: Mapped[str] = mapped_column(String(255), default="")
    engagement_type: Mapped[str] = mapped_column(String(255), default="")

    assessment_start: Mapped[str] = mapped_column(String(50), default="")
    assessment_end: Mapped[str] = mapped_column(String(50), default="")

    target_ips: Mapped[str] = mapped_column(Text, default="")
    target_domains: Mapped[str] = mapped_column(Text, default="")
    web_app_name: Mapped[str] = mapped_column(String(255), default="")
    mobile_app_name: Mapped[str] = mapped_column(String(255), default="")
    internal_app_name: Mapped[str] = mapped_column(String(255), default="")
    environment_name: Mapped[str] = mapped_column(String(255), default="")

    scope_summary: Mapped[str] = mapped_column(Text, default="")
    out_of_scope: Mapped[str] = mapped_column(Text, default="")
    attack_surface: Mapped[str] = mapped_column(Text, default="")
    methodology_summary: Mapped[str] = mapped_column(Text, default="")
    standards_used: Mapped[str] = mapped_column(Text, default="")

    stakeholder_summary: Mapped[str] = mapped_column(Text, default="")
    executive_summary: Mapped[str] = mapped_column(Text, default="")
    technical_summary: Mapped[str] = mapped_column(Text, default="")

    project_notes: Mapped[str] = mapped_column(Text, default="")
    conclusions: Mapped[str] = mapped_column(Text, default="")
    risk_rating: Mapped[str] = mapped_column(String(50), default="Not Rated")

    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)