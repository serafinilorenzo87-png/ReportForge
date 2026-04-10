import json

from sqlalchemy import ForeignKey, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.services.database import Base


class Finding(Base):
    __tablename__ = "findings"

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id"), nullable=False)

    title: Mapped[str] = mapped_column(String(255), nullable=False)
    severity: Mapped[str] = mapped_column(String(50), default="Info")
    description: Mapped[str] = mapped_column(Text, default="")
    evidence: Mapped[str] = mapped_column(Text, default="")
    remediation: Mapped[str] = mapped_column(Text, default="")
    attachments_json: Mapped[str] = mapped_column(Text, default="[]")

    def get_attachments(self) -> list[str]:
        try:
            data = json.loads(self.attachments_json or "[]")
            if isinstance(data, list):
                return [str(item) for item in data]
            return []
        except Exception:
            return []

    def set_attachments(self, attachments: list[str]) -> None:
        self.attachments_json = json.dumps(attachments, ensure_ascii=False)