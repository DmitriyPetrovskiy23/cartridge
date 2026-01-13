from sqlalchemy import Column, Integer, String, Date, DateTime, ForeignKey, Text
from sqlalchemy.orm import relationship
from datetime import datetime, date
from src.database import Base


class Department(Base):
    __tablename__ = "departments"
    
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    manager = Column(String(100))
    phone = Column(String(20))
    employee_count = Column(Integer, default=0)
    
    employees = relationship("Employee", back_populates="department", cascade="all, delete-orphan")


class Employee(Base):
    __tablename__ = "employees"
    
    id = Column(Integer, primary_key=True, index=True)
    full_name = Column(String(150), nullable=False)
    position = Column(String(100))
    department_id = Column(Integer, ForeignKey("departments.id"))
    personnel_number = Column(String(20))
    phone = Column(String(20))
    email = Column(String(100))
    
    department = relationship("Department", back_populates="employees")


class Cartridge(Base):
    __tablename__ = "cartridges"
    
    id = Column(Integer, primary_key=True, index=True)
    article = Column(String(50), nullable=False)
    model = Column(String(100), nullable=False)
    printer_type = Column(String(50))
    color = Column(String(30))
    status = Column(String(30), default="новый")
    capacity = Column(Integer)
    initial_quantity = Column(Integer, default=0)
    total_quantity = Column(Integer, default=0)
    production_date = Column(Date)
    warranty_months = Column(Integer, default=12)
    
    locations = relationship("CartridgeLocation", back_populates="cartridge", cascade="all, delete-orphan")
    service_notes = relationship("ServiceNote", back_populates="cartridge")
    movements = relationship("CartridgeMovement", back_populates="cartridge")


class Warehouse(Base):
    __tablename__ = "warehouses"
    
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    location = Column(String(200))
    description = Column(Text)
    
    boxes = relationship("Box", back_populates="warehouse", cascade="all, delete-orphan")


class Box(Base):
    __tablename__ = "boxes"
    
    id = Column(Integer, primary_key=True, index=True)
    warehouse_id = Column(Integer, ForeignKey("warehouses.id"))
    box_number = Column(String(20), nullable=False)
    description = Column(String(200))
    capacity = Column(Integer, default=10)
    current_count = Column(Integer, default=0)
    
    warehouse = relationship("Warehouse", back_populates="boxes")
    locations = relationship("CartridgeLocation", back_populates="box", cascade="all, delete-orphan")


class CartridgeLocation(Base):
    __tablename__ = "cartridge_locations"
    
    id = Column(Integer, primary_key=True, index=True)
    cartridge_id = Column(Integer, ForeignKey("cartridges.id"))
    box_id = Column(Integer, ForeignKey("boxes.id"), nullable=True)
    employee_id = Column(Integer, ForeignKey("employees.id"), nullable=True)
    status = Column(String(30), default="на складе")
    placed_date = Column(DateTime, default=datetime.utcnow)
    quantity = Column(Integer, default=1)
    
    cartridge = relationship("Cartridge", back_populates="locations")
    box = relationship("Box", back_populates="locations")
    employee = relationship("Employee")


class ServiceNote(Base):
    __tablename__ = "service_notes"
    
    id = Column(Integer, primary_key=True, index=True)
    note_number = Column(String(20), unique=True, nullable=False)
    created_date = Column(DateTime, default=datetime.utcnow)
    author_id = Column(Integer, ForeignKey("employees.id"))
    recipient_id = Column(Integer, ForeignKey("employees.id"))
    cartridge_id = Column(Integer, ForeignKey("cartridges.id"))
    quantity = Column(Integer, default=1)
    box_id = Column(Integer, ForeignKey("boxes.id"))
    reason = Column(String(50))
    comment = Column(Text)
    status = Column(String(30), default="запрошено")
    
    author = relationship("Employee", foreign_keys=[author_id])
    recipient = relationship("Employee", foreign_keys=[recipient_id])
    cartridge = relationship("Cartridge", back_populates="service_notes")
    box = relationship("Box")


class CartridgeMovement(Base):
    __tablename__ = "cartridge_movements"
    
    id = Column(Integer, primary_key=True, index=True)
    cartridge_id = Column(Integer, ForeignKey("cartridges.id"))
    from_location = Column(String(100))
    to_location = Column(String(100))
    movement_date = Column(DateTime, default=datetime.utcnow)
    service_note_id = Column(Integer, ForeignKey("service_notes.id"), nullable=True)
    
    cartridge = relationship("Cartridge", back_populates="movements")
    service_note = relationship("ServiceNote")
