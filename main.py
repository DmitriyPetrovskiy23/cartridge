from fastapi import FastAPI, Depends, HTTPException, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from sqlalchemy import func
from datetime import datetime
from typing import Optional
from docx import Document
from docx.shared import Pt, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import os
import tempfile

from src.database import engine, get_db, Base
from src.models import (
    Department, Employee, Cartridge, Warehouse, Box, 
    CartridgeLocation, ServiceNote, CartridgeMovement
)
from src.schemas import (
    DepartmentCreate, EmployeeCreate, CartridgeCreate, 
    WarehouseCreate, BoxCreate, ServiceNoteCreate,
    CartridgeLocationCreate
)

Base.metadata.create_all(bind=engine)

app = FastAPI(title="Учет картриджей")

templates = Jinja2Templates(directory="templates")


def generate_note_number(db: Session) -> str:
    year = datetime.now().year
    count = db.query(ServiceNote).filter(
        func.extract('year', ServiceNote.created_date) == year
    ).count()
    return f"КАРТ-{year}-{str(count + 1).zfill(3)}"


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request, db: Session = Depends(get_db)):
    total_cartridges = db.query(func.sum(CartridgeLocation.quantity)).filter(
        CartridgeLocation.status == "на складе"
    ).scalar() or 0
    
    in_use = db.query(func.sum(CartridgeLocation.quantity)).filter(
        CartridgeLocation.status == "выдано"
    ).scalar() or 0
    
    recent_notes = db.query(ServiceNote).order_by(
        ServiceNote.created_date.desc()
    ).limit(5).all()
    
    low_stock_boxes = db.query(Box).filter(Box.current_count < 3).all()
    
    return templates.TemplateResponse("dashboard.html", {
        "request": request,
        "total_cartridges": total_cartridges,
        "in_use": in_use,
        "recent_notes": recent_notes,
        "low_stock_boxes": low_stock_boxes
    })


@app.get("/cartridges", response_class=HTMLResponse)
async def cartridges_page(request: Request, db: Session = Depends(get_db)):
    cartridges = db.query(Cartridge).all()
    return templates.TemplateResponse("cartridges.html", {
        "request": request, 
        "cartridges": cartridges
    })


@app.post("/cartridges")
async def create_cartridge(
    request: Request,
    article: str = Form(...),
    model: str = Form(...),
    printer_type: str = Form(None),
    color: str = Form(None),
    status: str = Form("новый"),
    capacity: int = Form(None),
    production_date: str = Form(None),
    warranty_months: int = Form(12),
    db: Session = Depends(get_db)
):
    prod_date = None
    if production_date:
        prod_date = datetime.strptime(production_date, "%Y-%m-%d").date()
    
    cartridge = Cartridge(
        article=article,
        model=model,
        printer_type=printer_type,
        color=color,
        status=status,
        capacity=capacity,
        production_date=prod_date,
        warranty_months=warranty_months
    )
    db.add(cartridge)
    db.commit()
    return RedirectResponse(url="/cartridges", status_code=303)


@app.get("/warehouses", response_class=HTMLResponse)
async def warehouses_page(request: Request, db: Session = Depends(get_db)):
    warehouses = db.query(Warehouse).all()
    boxes = db.query(Box).all()
    cartridges = db.query(Cartridge).all()
    locations = db.query(CartridgeLocation).filter(
        CartridgeLocation.status == "на складе"
    ).all()
    return templates.TemplateResponse("warehouses.html", {
        "request": request, 
        "warehouses": warehouses,
        "boxes": boxes,
        "cartridges": cartridges,
        "locations": locations
    })


@app.post("/warehouses")
async def create_warehouse(
    name: str = Form(...),
    location: str = Form(None),
    description: str = Form(None),
    db: Session = Depends(get_db)
):
    warehouse = Warehouse(name=name, location=location, description=description)
    db.add(warehouse)
    db.commit()
    return RedirectResponse(url="/warehouses", status_code=303)


@app.post("/boxes")
async def create_box(
    warehouse_id: int = Form(...),
    box_number: str = Form(...),
    description: str = Form(None),
    capacity: int = Form(10),
    db: Session = Depends(get_db)
):
    box = Box(
        warehouse_id=warehouse_id,
        box_number=box_number,
        description=description,
        capacity=capacity
    )
    db.add(box)
    db.commit()
    return RedirectResponse(url="/warehouses", status_code=303)


@app.get("/departments", response_class=HTMLResponse)
async def departments_page(request: Request, db: Session = Depends(get_db)):
    departments = db.query(Department).all()
    return templates.TemplateResponse("departments.html", {
        "request": request, 
        "departments": departments
    })


@app.post("/departments")
async def create_department(
    name: str = Form(...),
    manager: str = Form(None),
    phone: str = Form(None),
    employee_count: int = Form(0),
    db: Session = Depends(get_db)
):
    department = Department(
        name=name, 
        manager=manager, 
        phone=phone, 
        employee_count=employee_count
    )
    db.add(department)
    db.commit()
    return RedirectResponse(url="/departments", status_code=303)


@app.get("/employees", response_class=HTMLResponse)
async def employees_page(request: Request, db: Session = Depends(get_db)):
    employees = db.query(Employee).all()
    departments = db.query(Department).all()
    return templates.TemplateResponse("employees.html", {
        "request": request, 
        "employees": employees,
        "departments": departments
    })


@app.post("/employees")
async def create_employee(
    full_name: str = Form(...),
    position: str = Form(None),
    department_id: int = Form(None),
    personnel_number: str = Form(None),
    phone: str = Form(None),
    email: str = Form(None),
    db: Session = Depends(get_db)
):
    employee = Employee(
        full_name=full_name,
        position=position,
        department_id=department_id,
        personnel_number=personnel_number,
        phone=phone,
        email=email
    )
    db.add(employee)
    db.commit()
    return RedirectResponse(url="/employees", status_code=303)


@app.get("/service-notes", response_class=HTMLResponse)
async def service_notes_page(request: Request, db: Session = Depends(get_db)):
    notes = db.query(ServiceNote).order_by(ServiceNote.created_date.desc()).all()
    employees = db.query(Employee).all()
    cartridges = db.query(Cartridge).all()
    boxes = db.query(Box).filter(Box.current_count > 0).all()
    locations = db.query(CartridgeLocation).filter(
        CartridgeLocation.status == "на складе",
        CartridgeLocation.quantity > 0
    ).all()
    
    return templates.TemplateResponse("service_notes.html", {
        "request": request, 
        "notes": notes,
        "employees": employees,
        "cartridges": cartridges,
        "boxes": boxes,
        "locations": locations
    })


@app.post("/service-notes")
async def create_service_note(
    author_id: int = Form(...),
    recipient_id: int = Form(...),
    cartridge_id: int = Form(...),
    quantity: int = Form(1),
    box_id: int = Form(...),
    reason: str = Form(...),
    comment: str = Form(None),
    db: Session = Depends(get_db)
):
    location = db.query(CartridgeLocation).filter(
        CartridgeLocation.cartridge_id == cartridge_id,
        CartridgeLocation.box_id == box_id,
        CartridgeLocation.status == "на складе"
    ).first()
    
    if not location or location.quantity < quantity:
        raise HTTPException(status_code=400, detail="Недостаточно картриджей на складе")
    
    note_number = generate_note_number(db)
    
    note = ServiceNote(
        note_number=note_number,
        author_id=author_id,
        recipient_id=recipient_id,
        cartridge_id=cartridge_id,
        quantity=quantity,
        box_id=box_id,
        reason=reason,
        comment=comment,
        status="выдано"
    )
    db.add(note)
    
    location.quantity -= quantity
    if location.quantity <= 0:
        location.status = "выдано"
    
    box = db.query(Box).filter(Box.id == box_id).first()
    if box:
        box.current_count = max(0, box.current_count - quantity)
    
    recipient = db.query(Employee).filter(Employee.id == recipient_id).first()
    cartridge = db.query(Cartridge).filter(Cartridge.id == cartridge_id).first()
    
    movement = CartridgeMovement(
        cartridge_id=cartridge_id,
        from_location=f"Ящик {box.box_number}" if box else "Склад",
        to_location=f"Сотрудник: {recipient.full_name}" if recipient else "Выдано",
        service_note_id=note.id
    )
    db.add(movement)
    
    db.commit()
    return RedirectResponse(url="/service-notes", status_code=303)


@app.post("/service-notes/{note_id}/return")
async def return_cartridge(note_id: int, db: Session = Depends(get_db)):
    note = db.query(ServiceNote).filter(ServiceNote.id == note_id).first()
    if not note:
        raise HTTPException(status_code=404, detail="Служебка не найдена")
    
    note.status = "возврат"
    
    location = db.query(CartridgeLocation).filter(
        CartridgeLocation.cartridge_id == note.cartridge_id,
        CartridgeLocation.box_id == note.box_id
    ).first()
    
    if location:
        location.quantity += note.quantity
        location.status = "на складе"
    else:
        new_location = CartridgeLocation(
            cartridge_id=note.cartridge_id,
            box_id=note.box_id,
            status="на складе",
            quantity=note.quantity
        )
        db.add(new_location)
    
    box = db.query(Box).filter(Box.id == note.box_id).first()
    if box:
        box.current_count += note.quantity
    
    movement = CartridgeMovement(
        cartridge_id=note.cartridge_id,
        from_location="В использовании",
        to_location=f"Ящик {box.box_number}" if box else "Склад",
        service_note_id=note.id
    )
    db.add(movement)
    
    db.commit()
    return RedirectResponse(url="/service-notes", status_code=303)


@app.post("/locations")
async def add_cartridge_to_stock(
    cartridge_id: int = Form(...),
    box_id: int = Form(...),
    quantity: int = Form(1),
    db: Session = Depends(get_db)
):
    location = db.query(CartridgeLocation).filter(
        CartridgeLocation.cartridge_id == cartridge_id,
        CartridgeLocation.box_id == box_id,
        CartridgeLocation.status == "на складе"
    ).first()
    
    if location:
        location.quantity += quantity
    else:
        location = CartridgeLocation(
            cartridge_id=cartridge_id,
            box_id=box_id,
            status="на складе",
            quantity=quantity
        )
        db.add(location)
    
    box = db.query(Box).filter(Box.id == box_id).first()
    if box:
        box.current_count += quantity
    
    movement = CartridgeMovement(
        cartridge_id=cartridge_id,
        from_location="Поступление",
        to_location=f"Ящик {box.box_number}" if box else "Склад"
    )
    db.add(movement)
    
    db.commit()
    return RedirectResponse(url="/warehouses", status_code=303)


@app.get("/reports", response_class=HTMLResponse)
async def reports_page(request: Request, db: Session = Depends(get_db)):
    inventory = db.query(
        Box.box_number,
        Warehouse.name.label("warehouse_name"),
        func.sum(CartridgeLocation.quantity).label("total")
    ).join(Warehouse).outerjoin(CartridgeLocation).filter(
        CartridgeLocation.status == "на складе"
    ).group_by(Box.id, Warehouse.name).all()
    
    dept_stats = db.query(
        Department.name,
        func.count(ServiceNote.id).label("notes_count")
    ).join(Employee, Employee.department_id == Department.id).join(
        ServiceNote, ServiceNote.recipient_id == Employee.id
    ).group_by(Department.id).all()
    
    low_stock = db.query(Box).filter(Box.current_count < 3).all()
    
    movements = db.query(CartridgeMovement).order_by(
        CartridgeMovement.movement_date.desc()
    ).limit(20).all()
    
    return templates.TemplateResponse("reports.html", {
        "request": request,
        "inventory": inventory,
        "dept_stats": dept_stats,
        "low_stock": low_stock,
        "movements": movements
    })


@app.get("/api/employee/{employee_id}/department")
async def get_employee_department(employee_id: int, db: Session = Depends(get_db)):
    employee = db.query(Employee).filter(Employee.id == employee_id).first()
    if employee and employee.department:
        return {"department_name": employee.department.name}
    return {"department_name": ""}


@app.get("/service-notes/{note_id}/print")
async def print_service_note(note_id: int, db: Session = Depends(get_db)):
    note = db.query(ServiceNote).filter(ServiceNote.id == note_id).first()
    if not note:
        raise HTTPException(status_code=404, detail="Служебка не найдена")
    
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    
    author_position = note.author.position if note.author and note.author.position else ""
    if note.author and note.author.full_name:
        name_parts = note.author.full_name.split()
        if len(name_parts) >= 3:
            author_short = f"{name_parts[0]} {name_parts[1][0]}. {name_parts[2][0]}."
        elif len(name_parts) == 2:
            author_short = f"{name_parts[0]} {name_parts[1][0]}."
        else:
            author_short = note.author.full_name
    else:
        author_short = "_______________"
    
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    header.paragraph_format.left_indent = Cm(8.25)
    header.paragraph_format.right_indent = Cm(-1.76)
    header.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    header.add_run("Исполняющему обязанности начальника\nотдела информатизации\nКраснодарского филиала\nРЭУ им. Г.В. Плеханова\nПетровскому Д. А.\n")
    if author_position:
        header.add_run(f"{author_position}\n")
    header.add_run(f"{author_short}")
    
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("служебная записка.")
    run.font.size = Pt(14)
    
    doc.add_paragraph()
    
    cartridge_info = f"{note.cartridge.article}" if note.cartridge else "картридж"
    cartridge_model = f"{note.cartridge.model}" if note.cartridge else ""
    printer_type = note.cartridge.printer_type if note.cartridge and note.cartridge.printer_type else "лазерного"
    
    body = doc.add_paragraph()
    body.alignment = WD_ALIGN_PARAGRAPH.CENTER
    body.add_run(f"Прошу предоставить картридж {cartridge_info} для лазерного принтера {cartridge_model} в количестве {note.quantity} шт.")
    
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    
    signature = doc.add_paragraph()
    signature.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    signature.add_run("_______________________")
    
    temp_dir = tempfile.gettempdir()
    filename = f"service_note_{note.note_number.replace('-', '_')}.docx"
    filepath = os.path.join(temp_dir, filename)
    doc.save(filepath)
    
    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000)
