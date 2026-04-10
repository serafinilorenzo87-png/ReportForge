from app.models.checklist_item import ChecklistItem
from app.models.finding import Finding
from app.models.project import Project
from app.models.reference import Reference
from app.services.database import Base, engine


def init_database() -> None:
    Base.metadata.create_all(bind=engine)