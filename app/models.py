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
class User(Base):
    __tablename__='users'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); username=Column(String(80),unique=True,nullable=False); full_name=Column(String(150),nullable=False); role=Column(String(80),default='مستخدم'); permissions=Column(String(500),default=''); is_active=Column(Boolean,default=True,nullable=False)
class Customer(Base):
    __tablename__='customers'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(160),unique=True,nullable=False); phone=Column(String(50),default=''); opening_debit=Column(Float,default=0,nullable=False); opening_credit=Column(Float,default=0,nullable=False); is_active=Column(Boolean,default=True,nullable=False)
class Treasury(Base):
    __tablename__='treasuries'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(150),unique=True,nullable=False); branch_id=Column(Integer,ForeignKey('branches.id'),nullable=False); opening_balance=Column(Float,default=0,nullable=False); is_active=Column(Boolean,default=True,nullable=False)
    branch=relationship('Branch')
class Bank(Base):
    __tablename__='banks'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(150),nullable=False); account_number=Column(String(100),default=''); opening_balance=Column(Float,default=0,nullable=False); is_active=Column(Boolean,default=True,nullable=False)
class ExpenseItem(Base):
    __tablename__='expense_items'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(150),unique=True,nullable=False); description=Column(String(300),default=''); expense_type=Column(String(20),default='operating',nullable=False); is_active=Column(Boolean,default=True,nullable=False)
class OtherAccountItem(Base):
    __tablename__='other_account_items'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); name=Column(String(150),unique=True,nullable=False); account_type=Column(String(80),default=''); effect_sign=Column(Integer,default=1,nullable=False); description=Column(String(300),default=''); opening_debit=Column(Float,default=0,nullable=False); opening_credit=Column(Float,default=0,nullable=False); is_active=Column(Boolean,default=True,nullable=False)
class OpeningStock(Base):
    __tablename__='opening_stocks'
    id=Column(Integer,primary_key=True); code=Column(String(20),unique=True,nullable=False); item_name=Column(String(180),nullable=False); branch_id=Column(Integer,ForeignKey('branches.id'),nullable=False); quantity=Column(Float,default=0,nullable=False); unit_cost=Column(Float,default=0,nullable=False); notes=Column(String(300),default='')
    branch=relationship('Branch')
class TreasuryDeposit(Base):
    __tablename__='treasury_deposits'
    id=Column(Integer,primary_key=True); sales_journal_id=Column(Integer,ForeignKey('sales_journals.id'),unique=True,nullable=False); treasury_id=Column(Integer,ForeignKey('treasuries.id'),nullable=False); amount=Column(Float,default=0,nullable=False); created_at=Column(DateTime,default=datetime.utcnow,nullable=False)
    treasury=relationship('Treasury'); sales_journal=relationship('SalesJournal',back_populates='treasury_deposit')
class SalesJournal(Base):
    __tablename__='sales_journals'; __table_args__=(UniqueConstraint('branch_id','journal_date',name='uq_sales_branch_date'),)
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,default=date.today,index=True); branch_id=Column(Integer,ForeignKey('branches.id'),nullable=False,index=True); status=Column(String(20),default='draft',nullable=False,index=True); notes=Column(String(500),default=''); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); updated_at=Column(DateTime,default=datetime.utcnow,onupdate=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    branch=relationship('Branch',back_populates='sales_journals'); lines=relationship('SalesLine',back_populates='journal',cascade='all, delete-orphan',order_by='SalesLine.id'); treasury_deposit=relationship('TreasuryDeposit',back_populates='sales_journal',uselist=False,cascade='all, delete-orphan')
    total_shift=property(lambda s:sum(x.shift_value or 0 for x in s.lines)); total_discount=property(lambda s:sum(x.discount or 0 for x in s.lines)); total_net_cash=property(lambda s:sum(x.net_cash or 0 for x in s.lines)); total_cash_difference=property(lambda s:sum(x.cash_difference or 0 for x in s.lines))
class SalesLine(Base):
    __tablename__='sales_lines'
    id=Column(Integer,primary_key=True); journal_id=Column(Integer,ForeignKey('sales_journals.id'),nullable=False,index=True); employee_id=Column(Integer,ForeignKey('employees.id'),nullable=False,index=True); shift_value=Column(Float,default=0,nullable=False); discount=Column(Float,default=0,nullable=False); net_cash=Column(Float,default=0,nullable=False); cash_difference=Column(Float,default=0,nullable=False)
    journal=relationship('SalesJournal',back_populates='lines'); employee=relationship('Employee',back_populates='sales_lines')
class PurchaseJournal(Base):
    __tablename__='purchase_journals'
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,index=True); entry_type=Column(String(20),nullable=False,default='purchase'); notice_type=Column(String(40),default='',nullable=False); status=Column(String(20),default='draft',nullable=False,index=True); notes=Column(String(500),default=''); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    lines=relationship('PurchaseLine',back_populates='journal',cascade='all, delete-orphan',order_by='PurchaseLine.id')
    total_pharmacy=property(lambda s:sum(x.pharmacy_value or 0 for x in s.lines)); total_public=property(lambda s:sum(x.public_value or 0 for x in s.lines)); total_effect=property(lambda s:sum(x.account_effect or 0 for x in s.lines))
class PurchaseLine(Base):
    __tablename__='purchase_lines'; __table_args__=(UniqueConstraint('supplier_id','document_no','entry_type',name='uq_supplier_document_type'),)
    id=Column(Integer,primary_key=True); journal_id=Column(Integer,ForeignKey('purchase_journals.id'),nullable=False,index=True); supplier_id=Column(Integer,ForeignKey('suppliers.id'),nullable=False,index=True); entry_type=Column(String(20),nullable=False); document_no=Column(String(80),nullable=False,index=True); notice_type=Column(String(40),default=''); pharmacy_value=Column(Float,default=0,nullable=False); public_value=Column(Float,default=0,nullable=False); discount_percent=Column(Float,default=0,nullable=False); account_effect=Column(Float,default=0,nullable=False); description=Column(String(250),default=''); notes=Column(String(500),default='')
    journal=relationship('PurchaseJournal',back_populates='lines'); supplier=relationship('Supplier',back_populates='purchase_lines')
class ExpenseJournal(Base):
    __tablename__='expense_journals'
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,index=True); expense_type=Column(String(20),nullable=False,index=True); branch_id=Column(Integer,ForeignKey('branches.id'),index=True); status=Column(String(20),default='draft',nullable=False,index=True); notes=Column(String(500),default=''); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); updated_at=Column(DateTime,default=datetime.utcnow,onupdate=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    branch=relationship('Branch'); lines=relationship('ExpenseLine',back_populates='journal',cascade='all, delete-orphan',order_by='ExpenseLine.id'); treasury_payment=relationship('ExpenseTreasuryPayment',back_populates='journal',uselist=False,cascade='all, delete-orphan'); total_amount=property(lambda s:sum(x.amount or 0 for x in s.lines))
class ExpenseLine(Base):
    __tablename__='expense_lines'
    id=Column(Integer,primary_key=True); journal_id=Column(Integer,ForeignKey('expense_journals.id'),nullable=False,index=True); expense_item_id=Column(Integer,ForeignKey('expense_items.id'),nullable=False,index=True); amount=Column(Float,default=0,nullable=False); notes=Column(String(500),default='')
    journal=relationship('ExpenseJournal',back_populates='lines'); expense_item=relationship('ExpenseItem')
class ExpenseTreasuryPayment(Base):
    __tablename__='expense_treasury_payments'
    id=Column(Integer,primary_key=True); expense_journal_id=Column(Integer,ForeignKey('expense_journals.id'),unique=True,nullable=False); treasury_id=Column(Integer,ForeignKey('treasuries.id'),nullable=False); amount=Column(Float,default=0,nullable=False); created_at=Column(DateTime,default=datetime.utcnow,nullable=False)
    journal=relationship('ExpenseJournal',back_populates='treasury_payment'); treasury=relationship('Treasury')
class OtherAccountJournal(Base):
    __tablename__='other_account_journals'
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,index=True); transaction_type=Column(String(20),nullable=False,index=True); status=Column(String(20),default='draft',nullable=False,index=True); notes=Column(String(500),default=''); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); updated_at=Column(DateTime,default=datetime.utcnow,onupdate=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    lines=relationship('OtherAccountLine',back_populates='journal',cascade='all, delete-orphan',order_by='OtherAccountLine.id'); total_amount=property(lambda s:sum(x.amount or 0 for x in s.lines))
class OtherAccountLine(Base):
    __tablename__='other_account_lines'
    id=Column(Integer,primary_key=True); journal_id=Column(Integer,ForeignKey('other_account_journals.id'),nullable=False,index=True); account_id=Column(Integer,ForeignKey('other_account_items.id'),nullable=False,index=True); amount=Column(Float,default=0,nullable=False); description=Column(String(500),default='')
    journal=relationship('OtherAccountJournal',back_populates='lines'); account=relationship('OtherAccountItem')
class SupplierClaim(Base):
    __tablename__='supplier_claims'
    id=Column(Integer,primary_key=True); claim_no=Column(String(30),unique=True,nullable=False,index=True); claim_date=Column(Date,nullable=False,index=True); supplier_id=Column(Integer,ForeignKey('suppliers.id'),nullable=False,index=True); status=Column(String(20),default='ready',nullable=False,index=True); total_amount=Column(Float,default=0,nullable=False); paid_amount=Column(Float,default=0,nullable=False); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); closed_at=Column(DateTime)
    supplier=relationship('Supplier'); lines=relationship('SupplierClaimLine',back_populates='claim',cascade='all, delete-orphan'); allocations=relationship('SupplierPaymentAllocation',back_populates='claim')
    remaining=property(lambda s:max((s.total_amount or 0)-(s.paid_amount or 0),0))
class SupplierClaimLine(Base):
    __tablename__='supplier_claim_lines'; __table_args__=(UniqueConstraint('purchase_line_id',name='uq_claim_purchase_line'),)
    id=Column(Integer,primary_key=True); claim_id=Column(Integer,ForeignKey('supplier_claims.id'),nullable=False,index=True); purchase_line_id=Column(Integer,ForeignKey('purchase_lines.id'),nullable=False,index=True); amount=Column(Float,default=0,nullable=False)
    claim=relationship('SupplierClaim',back_populates='lines'); purchase_line=relationship('PurchaseLine')
class SupplierPaymentJournal(Base):
    __tablename__='supplier_payment_journals'
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,index=True); supplier_id=Column(Integer,ForeignKey('suppliers.id'),nullable=False,index=True); payment_method=Column(String(20),nullable=False,index=True); treasury_id=Column(Integer,ForeignKey('treasuries.id'),index=True); total_amount=Column(Float,default=0,nullable=False); status=Column(String(20),default='draft',nullable=False,index=True); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    supplier=relationship('Supplier'); treasury=relationship('Treasury'); allocations=relationship('SupplierPaymentAllocation',back_populates='journal',cascade='all, delete-orphan'); checks=relationship('IssuedCheck',back_populates='payment_journal',cascade='all, delete-orphan')
    lines=property(lambda s:s.allocations); total_effect=property(lambda s:s.total_amount)
class SupplierPaymentAllocation(Base):
    __tablename__='supplier_payment_allocations'; __table_args__=(UniqueConstraint('journal_id','claim_id',name='uq_payment_claim'),)
    id=Column(Integer,primary_key=True); journal_id=Column(Integer,ForeignKey('supplier_payment_journals.id'),nullable=False,index=True); claim_id=Column(Integer,ForeignKey('supplier_claims.id'),nullable=False,index=True); amount=Column(Float,default=0,nullable=False)
    journal=relationship('SupplierPaymentJournal',back_populates='allocations'); claim=relationship('SupplierClaim',back_populates='allocations')
class GeneralCheckJournal(Base):
    __tablename__='general_check_journals'
    id=Column(Integer,primary_key=True); journal_no=Column(String(30),unique=True,nullable=False,index=True); journal_date=Column(Date,nullable=False,index=True); account_id=Column(Integer,ForeignKey('other_account_items.id'),nullable=False,index=True); status=Column(String(20),default='draft',nullable=False,index=True); created_at=Column(DateTime,default=datetime.utcnow,nullable=False); posted_at=Column(DateTime)
    account=relationship('OtherAccountItem'); checks=relationship('IssuedCheck',back_populates='general_journal',cascade='all, delete-orphan'); lines=property(lambda s:s.checks); total_amount=property(lambda s:sum(x.amount or 0 for x in s.checks)); total_effect=property(lambda s:s.total_amount)
class IssuedCheck(Base):
    __tablename__='issued_checks'; __table_args__=(UniqueConstraint('bank_id','check_no',name='uq_bank_check_no'),)
    id=Column(Integer,primary_key=True); payment_journal_id=Column(Integer,ForeignKey('supplier_payment_journals.id'),index=True); general_journal_id=Column(Integer,ForeignKey('general_check_journals.id'),index=True); supplier_id=Column(Integer,ForeignKey('suppliers.id'),index=True); general_account_id=Column(Integer,ForeignKey('other_account_items.id'),index=True); bank_id=Column(Integer,ForeignKey('banks.id'),nullable=False,index=True); check_no=Column(String(80),nullable=False,index=True); amount=Column(Float,default=0,nullable=False); due_date=Column(Date,nullable=False,index=True); beneficiary=Column(String(180),nullable=False); description=Column(String(500),default=''); status=Column(String(20),default='draft',nullable=False,index=True); posted_at=Column(DateTime); cleared_at=Column(DateTime); notification_key=Column(String(100),default='')
    payment_journal=relationship('SupplierPaymentJournal',back_populates='checks'); general_journal=relationship('GeneralCheckJournal',back_populates='checks'); supplier=relationship('Supplier'); general_account=relationship('OtherAccountItem'); bank=relationship('Bank')
class NotificationRead(Base):
    __tablename__='notification_reads'
    id=Column(Integer,primary_key=True); notification_key=Column(String(120),unique=True,nullable=False,index=True); read_at=Column(DateTime,default=datetime.utcnow,nullable=False)
