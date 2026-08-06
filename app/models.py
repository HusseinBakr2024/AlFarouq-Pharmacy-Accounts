from datetime import date, datetime
from sqlalchemy import Boolean, Column, Date, DateTime, Float, ForeignKey, Integer, String, UniqueConstraint
from sqlalchemy.orm import relationship
from app.database import Base

class Branch(Base):
    __tablename__='branches'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(120),unique=True,nullable=False); is_active=Column(Boolean,default=True,nullable=False)
    employees=relationship('Employee',back_populates='branch'); sales_journals=relationship('SalesJournal',back_populates='branch')
class Employee(Base):
    __tablename__='employees'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(150),nullable=False); job_title=Column(String(100),default=''); branch_id=Column(Integer,ForeignKey('branches.id'),nullable=False); is_active=Column(Boolean,default=True,nullable=False)
    branch=relationship('Branch',back_populates='employees'); sales_lines=relationship('SalesLine',back_populates='employee')
class Supplier(Base):
    __tablename__='suppliers'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(160),unique=True,nullable=False); phone=Column(String(50),default=''); opening_debit=Column(Float,default=0,nullable=False); opening_credit=Column(Float,default=0,nullable=False); is_active=Column(Boolean,default=True,nullable=False)
    purchase_lines=relationship('PurchaseLine',back_populates='supplier')
class SalesJournal(Base):
    __tablename__='sales_journals'; __table_args__=(UniqueConstraint('branch_id','journal_date',name='uq_sales_branch_date'),)
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,default=date.today,index=True); branch_id=Column(Integer,ForeignKey('branches.id'),nullable=False,index=True); status=Column(String(20),default='draft',nullable=False,index=True); notes=Column(String(500),default=''); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); updated_at=Column(DateTime,default=datetime.utcnow,onupdate=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    branch=relationship('Branch',back_populates='sales_journals'); lines=relationship('SalesLine',back_populates='journal',cascade='all, delete-orphan',order_by='SalesLine.id')
    total_shift=property(lambda s:sum(x.shift_value or 0 for x in s.lines)); total_discount=property(lambda s:sum(x.discount or 0 for x in s.lines)); total_net_cash=property(lambda s:sum(x.net_cash or 0 for x in s.lines)); total_cash_difference=property(lambda s:sum(x.cash_difference or 0 for x in s.lines))
class SalesLine(Base):
    __tablename__='sales_lines'
    id=Column(Integer,primary_key=True); journal_id=Column(Integer,ForeignKey('sales_journals.id'),nullable=False,index=True); employee_id=Column(Integer,ForeignKey('employees.id'),nullable=False,index=True); shift_value=Column(Float,default=0,nullable=False); discount=Column(Float,default=0,nullable=False); net_cash=Column(Float,default=0,nullable=False); cash_difference=Column(Float,default=0,nullable=False)
    journal=relationship('SalesJournal',back_populates='lines'); employee=relationship('Employee',back_populates='sales_lines')
class PurchaseJournal(Base):
    __tablename__='purchase_journals'
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,index=True); entry_type=Column(String(20),nullable=False,default='purchase'); status=Column(String(20),default='draft',nullable=False,index=True); notes=Column(String(500),default=''); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    lines=relationship('PurchaseLine',back_populates='journal',cascade='all, delete-orphan',order_by='PurchaseLine.id')
    total_pharmacy=property(lambda s:sum(x.pharmacy_value or 0 for x in s.lines)); total_public=property(lambda s:sum(x.public_value or 0 for x in s.lines)); total_effect=property(lambda s:sum(x.account_effect or 0 for x in s.lines))
class PurchaseLine(Base):
    __tablename__='purchase_lines'; __table_args__=(UniqueConstraint('supplier_id','document_no','entry_type',name='uq_supplier_document_type'),)
    id=Column(Integer,primary_key=True); journal_id=Column(Integer,ForeignKey('purchase_journals.id'),nullable=False,index=True); supplier_id=Column(Integer,ForeignKey('suppliers.id'),nullable=False,index=True); entry_type=Column(String(20),nullable=False); document_no=Column(String(80),nullable=False,index=True); notice_type=Column(String(40),default=''); pharmacy_value=Column(Float,default=0,nullable=False); public_value=Column(Float,default=0,nullable=False); discount_percent=Column(Float,default=0,nullable=False); account_effect=Column(Float,default=0,nullable=False); description=Column(String(250),default=''); notes=Column(String(500),default='')
    journal=relationship('PurchaseJournal',back_populates='lines'); supplier=relationship('Supplier',back_populates='purchase_lines')
