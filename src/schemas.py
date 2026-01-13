from pydantic import BaseModel
from datetime import date, datetime
from typing import Optional


class DepartmentCreate(BaseModel):
    name: str
    manager: Optional[str] = None
    phone: Optional[str] = None
    employee_count: int = 0


class DepartmentResponse(DepartmentCreate):
    id: int
    
    class Config:
        from_attributes = True


class EmployeeCreate(BaseModel):
    full_name: str
    position: Optional[str] = None
    department_id: Optional[int] = None
    personnel_number: Optional[str] = None
    phone: Optional[str] = None
    email: Optional[str] = None


class EmployeeResponse(EmployeeCreate):
    id: int
    
    class Config:
        from_attributes = True


class CartridgeCreate(BaseModel):
    article: str
    model: str
    printer_type: Optional[str] = None
    color: Optional[str] = None
    status: str = "новый"
    capacity: Optional[int] = None
    production_date: Optional[date] = None
    warranty_months: int = 12


class CartridgeResponse(CartridgeCreate):
    id: int
    
    class Config:
        from_attributes = True


class WarehouseCreate(BaseModel):
    name: str
    location: Optional[str] = None
    description: Optional[str] = None


class WarehouseResponse(WarehouseCreate):
    id: int
    
    class Config:
        from_attributes = True


class BoxCreate(BaseModel):
    warehouse_id: int
    box_number: str
    description: Optional[str] = None
    capacity: int = 10


class BoxResponse(BoxCreate):
    id: int
    current_count: int = 0
    
    class Config:
        from_attributes = True


class CartridgeLocationCreate(BaseModel):
    cartridge_id: int
    box_id: Optional[int] = None
    employee_id: Optional[int] = None
    status: str = "на складе"
    quantity: int = 1


class ServiceNoteCreate(BaseModel):
    author_id: int
    recipient_id: int
    cartridge_id: int
    quantity: int = 1
    box_id: int
    reason: str
    comment: Optional[str] = None


class ServiceNoteResponse(BaseModel):
    id: int
    note_number: str
    created_date: datetime
    author_id: int
    recipient_id: int
    cartridge_id: int
    quantity: int
    box_id: int
    reason: str
    comment: Optional[str]
    status: str
    
    class Config:
        from_attributes = True
