from datetime import date
from src.database import SessionLocal, engine, Base
from src.models import Department, Employee, Cartridge, Warehouse, Box, CartridgeLocation

Base.metadata.create_all(bind=engine)

def seed():
    db = SessionLocal()
    
    if db.query(Department).count() > 0:
        print("Данные уже существуют")
        db.close()
        return
    
    departments = [
        Department(name="Бухгалтерия", manager="Иванова А.П.", phone="+7 (495) 123-45-01", employee_count=5),
        Department(name="IT отдел", manager="Петров С.В.", phone="+7 (495) 123-45-02", employee_count=8),
        Department(name="Отдел продаж", manager="Сидорова М.К.", phone="+7 (495) 123-45-03", employee_count=12),
        Department(name="HR отдел", manager="Козлова Е.Н.", phone="+7 (495) 123-45-04", employee_count=3),
    ]
    db.add_all(departments)
    db.commit()
    
    employees = [
        Employee(full_name="Иванов Иван Иванович", position="Главный бухгалтер", department_id=1, personnel_number="001", phone="+7-900-111-1111", email="ivanov@company.ru"),
        Employee(full_name="Петрова Мария Сергеевна", position="Бухгалтер", department_id=1, personnel_number="002", phone="+7-900-222-2222", email="petrova@company.ru"),
        Employee(full_name="Сидоров Алексей Владимирович", position="Системный администратор", department_id=2, personnel_number="003", phone="+7-900-333-3333", email="sidorov@company.ru"),
        Employee(full_name="Козлова Елена Николаевна", position="Менеджер по продажам", department_id=3, personnel_number="004", phone="+7-900-444-4444", email="kozlova@company.ru"),
        Employee(full_name="Морозов Дмитрий Петрович", position="HR специалист", department_id=4, personnel_number="005", phone="+7-900-555-5555", email="morozov@company.ru"),
    ]
    db.add_all(employees)
    db.commit()
    
    cartridges = [
        Cartridge(article="HP-85A-BLK", model="HP LaserJet 85A", printer_type="лазерный", color="черный", status="новый", capacity=2000, production_date=date(2024, 1, 15), warranty_months=12),
        Cartridge(article="HP-83A-BLK", model="HP LaserJet 83A", printer_type="лазерный", color="черный", status="новый", capacity=1500, production_date=date(2024, 2, 10), warranty_months=12),
        Cartridge(article="CN-728-BLK", model="Canon 728", printer_type="лазерный", color="черный", status="новый", capacity=2100, production_date=date(2024, 3, 5), warranty_months=12),
        Cartridge(article="SM-D111S-BLK", model="Samsung MLT-D111S", printer_type="лазерный", color="черный", status="новый", capacity=1000, production_date=date(2024, 1, 20), warranty_months=12),
        Cartridge(article="HP-951XL-C", model="HP 951XL Cyan", printer_type="струйный", color="цветной", status="новый", capacity=1500, production_date=date(2024, 4, 1), warranty_months=6),
    ]
    db.add_all(cartridges)
    db.commit()
    
    warehouses = [
        Warehouse(name="Основной склад", location="Офис, 1 этаж", description="Главное хранилище расходных материалов"),
        Warehouse(name="Резервный склад", location="Офис, подвал", description="Дополнительное хранение"),
    ]
    db.add_all(warehouses)
    db.commit()
    
    boxes = [
        Box(warehouse_id=1, box_number="A-01", description="HP картриджи", capacity=20, current_count=0),
        Box(warehouse_id=1, box_number="A-02", description="Canon картриджи", capacity=15, current_count=0),
        Box(warehouse_id=1, box_number="B-01", description="Samsung картриджи", capacity=10, current_count=0),
        Box(warehouse_id=2, box_number="R-01", description="Резерв HP", capacity=10, current_count=0),
    ]
    db.add_all(boxes)
    db.commit()
    
    locations = [
        CartridgeLocation(cartridge_id=1, box_id=1, status="на складе", quantity=10),
        CartridgeLocation(cartridge_id=2, box_id=1, status="на складе", quantity=5),
        CartridgeLocation(cartridge_id=3, box_id=2, status="на складе", quantity=8),
        CartridgeLocation(cartridge_id=4, box_id=3, status="на складе", quantity=3),
        CartridgeLocation(cartridge_id=5, box_id=1, status="на складе", quantity=4),
    ]
    db.add_all(locations)
    
    boxes[0].current_count = 19
    boxes[1].current_count = 8
    boxes[2].current_count = 3
    
    db.commit()
    db.close()
    
    print("Демо-данные успешно добавлены!")

if __name__ == "__main__":
    seed()
