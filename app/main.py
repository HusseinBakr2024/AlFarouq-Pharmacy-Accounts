from datetime import date, datetime
from io import BytesIO
from typing import Optional

from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from sqlalchemy import func, or_, inspect, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session, joinedload

from app.database import Base, SessionLocal, engine, get_db
from app.models import (Branch, Employee, Supplier, SalesJournal, SalesLine, PurchaseJournal, PurchaseLine,
    User, Customer, Treasury, Bank, ExpenseItem, OtherAccountItem, OpeningStock, TreasuryDeposit,
    ExpenseJournal, ExpenseLine, ExpenseTreasuryPayment, OtherAccountJournal, OtherAccountLine)

Base.metadata.create_all(bind=engine)

# Keep existing SQLite installations compatible with additive master-data fields.
with engine.begin() as connection:
    columns={column["name"] for column in inspect(connection).get_columns("other_account_items")}
    if "effect_sign" not in columns:
        connection.execute(text("ALTER TABLE other_account_items ADD COLUMN effect_sign INTEGER NOT NULL DEFAULT 1"))
    purchase_journal_columns={column["name"] for column in inspect(connection).get_columns("purchase_journals")}
    if "notice_type" not in purchase_journal_columns:
        connection.execute(text("ALTER TABLE purchase_journals ADD COLUMN notice_type VARCHAR(40) NOT NULL DEFAULT ''"))
        connection.execute(text("UPDATE purchase_journals SET notice_type = COALESCE((SELECT notice_type FROM purchase_lines WHERE purchase_lines.journal_id = purchase_journals.id AND notice_type <> '' LIMIT 1), '') WHERE entry_type = 'notice'"))
    expense_columns={column["name"] for column in inspect(connection).get_columns("expense_items")}
    if "expense_type" not in expense_columns:
        connection.execute(text("ALTER TABLE expense_items ADD COLUMN expense_type VARCHAR(20) NOT NULL DEFAULT 'operating'"))
    account_columns={column["name"] for column in inspect(connection).get_columns("other_account_items")}
    if "opening_debit" not in account_columns:
        connection.execute(text("ALTER TABLE other_account_items ADD COLUMN opening_debit FLOAT NOT NULL DEFAULT 0"))
    if "opening_credit" not in account_columns:
        connection.execute(text("ALTER TABLE other_account_items ADD COLUMN opening_credit FLOAT NOT NULL DEFAULT 0"))

app = FastAPI(title="صيدليات الفاروق")
app.mount("/static", StaticFiles(directory="app/static"), name="static")
templates = Jinja2Templates(directory="app/templates")


def seed_master_data() -> None:
    db = SessionLocal()
    try:
        branches = db.query(Branch).order_by(Branch.id).all()
        if not branches:
            branches = [
                Branch(code="BR-0001", name="فرع الفاروق 1"),
                Branch(code="BR-0002", name="فرع الفاروق 2"),
            ]
            db.add_all(branches)
            db.flush()

        if db.query(Employee).count() == 0:
            first = branches[0]
            second = branches[1] if len(branches) > 1 else branches[0]
            db.add_all([
                Employee(code="EM-0001", name="مستخدم الوردية الأولى", job_title="كاشير", branch_id=first.id),
                Employee(code="EM-0002", name="مستخدم الوردية الثانية", job_title="كاشير", branch_id=first.id),
                Employee(code="EM-0003", name="مستخدم الوردية الأولى - فرع 2", job_title="كاشير", branch_id=second.id),
                Employee(code="EM-0004", name="مستخدم الوردية الثانية - فرع 2", job_title="كاشير", branch_id=second.id),
            ])
        db.commit()
    finally:
        db.close()


seed_master_data()


def render(request: Request, template_name: str, page_title: str, active_page: str, **context):
    context.update({"request": request, "page_title": page_title, "active_page": active_page})
    return templates.TemplateResponse(request=request, name=template_name, context=context)


def parse_date(value: Optional[str], default: Optional[date] = None) -> Optional[date]:
    if not value:
        return default
    value = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return default


def next_journal_no(db: Session, journal_date: date) -> str:
    prefix = f"SJ-{journal_date.year}-"
    last_no = db.query(func.max(SalesJournal.journal_no)).filter(SalesJournal.journal_no.like(f"{prefix}%")).scalar()
    sequence = int(last_no.rsplit("-", 1)[-1]) + 1 if last_no else 1
    return f"{prefix}{sequence:05d}"


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return render(request, "home/index.html", "الرئيسية", "home")


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    return render(request, "dashboard/index.html", "لوحة المعلومات", "dashboard")


def next_code(db: Session, model, prefix: str) -> str:
    last = db.query(model).order_by(model.id.desc()).first()
    return f"{prefix}-{((last.id if last else 0)+1):04d}"

SETTINGS = {
 "users": (User,"USR","المستخدمون والصلاحيات",[("username","اسم المستخدم","text"),("full_name","الاسم بالكامل","text"),("role","الدور","text"),("permissions","الصلاحيات","text"),("is_active","نشط","bool")]),
 "employees": (Employee,"EMP","الموظفون",[("name","اسم الموظف","text"),("branch_id","الفرع","branch"),("job_title","الوظيفة","text"),("is_active","نشط","bool")]),
 "branches": (Branch,"BR","الفروع",[("name","اسم الفرع","text"),("is_active","نشط","bool")]),
 "customers": (Customer,"CUS","العملاء",[("name","اسم العميل","text"),("phone","رقم التواصل","text"),("opening_debit","افتتاحي مدين","number"),("opening_credit","افتتاحي دائن","number"),("is_active","نشط","bool")]),
 "suppliers": (Supplier,"SUP","الموردون",[("name","اسم المورد","text"),("phone","رقم التواصل","text"),("opening_debit","افتتاحي مدين","number"),("opening_credit","افتتاحي دائن","number"),("is_active","نشط","bool")]),
 "treasuries": (Treasury,"TRE","الخزائن",[("name","اسم الخزينة","text"),("branch_id","الفرع","branch"),("opening_balance","الرصيد الافتتاحي","number"),("is_active","نشط","bool")]),
 "banks": (Bank,"BNK","البنوك",[("name","اسم البنك","text"),("account_number","رقم الحساب","text"),("opening_balance","الرصيد الافتتاحي","number"),("is_active","نشط","bool")]),
 "expenses": (ExpenseItem,"EXP","بنود المصروفات",[("name","اسم المصروف","text"),("expense_type","نوع المصروف","expense_type"),("is_active","نشط","bool")]),
 "other_accounts": (OtherAccountItem,"ACC","الحسابات الأخرى",[("name","اسم الحساب","text"),("opening_debit","افتتاحي مدين","number"),("opening_credit","افتتاحي دائن","number"),("is_active","نشط","bool")]),
 "opening_stock": (OpeningStock,"STK","المخزون الافتتاحي",[("item_name","اسم الصنف","text"),("branch_id","الفرع","branch"),("quantity","الكمية","number"),("unit_cost","تكلفة الوحدة","number"),("notes","ملاحظات","text")]),
}

@app.get("/settings", response_class=HTMLResponse)
def settings(request: Request, tab: str = "users", search: str = "", edit_id: Optional[int] = None, db: Session = Depends(get_db)):
    if tab not in SETTINGS: tab="users"
    model,prefix,title,fields=SETTINGS[tab]; query=db.query(model)
    searchable=[model.code]+[getattr(model,n) for n,_,kind in fields if kind=="text"]
    if search and searchable: query=query.filter(or_(*[x.ilike(f"%{search}%") for x in searchable]))
    records=query.order_by(model.id.desc()).all(); editing=db.get(model,edit_id) if edit_id else None
    tabs=[{"key":k,"title":v[2]} for k,v in SETTINGS.items()]
    return render(request,"settings/index.html","الإعدادات","settings",tab=tab,tabs=tabs,title=title,fields=fields,records=records,editing=editing,next_value=next_code(db,model,prefix),branches=db.query(Branch).filter(Branch.is_active.is_(True)).order_by(Branch.name).all(),search=search,message=request.query_params.get("message",""))

@app.post("/settings/{tab}/save")
async def save_setting(tab: str, request: Request, db: Session = Depends(get_db)):
    if tab not in SETTINGS: raise HTTPException(404)
    model,prefix,_,fields=SETTINGS[tab]; form=await request.form(); record=db.get(model,int(form.get("id"))) if form.get("id") else model(code=next_code(db,model,prefix))
    for name,_,kind in fields:
        raw=form.get(name)
        if kind=="bool": value=bool(raw)
        elif kind=="number": value=float(raw or 0)
        elif kind=="branch": value=int(raw or 0)
        else: value=str(raw or "").strip()
        setattr(record,name,value)
    if tab=="expenses" and record.expense_type not in {"operating","general"}:
        return RedirectResponse("/settings?tab=expenses&message=يجب اختيار نوع المصروف",303)
    if tab=="other_accounts":
        record.opening_debit=max(record.opening_debit,0); record.opening_credit=max(record.opening_credit,0)
        if record.opening_debit>0: record.opening_credit=0
    if not getattr(record,"id",None): db.add(record)
    try: db.commit()
    except IntegrityError:
        db.rollback(); return RedirectResponse(f"/settings?tab={tab}&message=تعذر الحفظ، تحقق من عدم تكرار البيانات",303)
    return RedirectResponse(f"/settings?tab={tab}&message=تم حفظ البيانات بنجاح",303)

@app.post("/settings/{tab}/{record_id}/delete")
def delete_setting(tab: str, record_id: int, db: Session = Depends(get_db)):
    if tab not in SETTINGS: raise HTTPException(404)
    record=db.get(SETTINGS[tab][0],record_id)
    if record:
        try: db.delete(record); db.commit()
        except IntegrityError: db.rollback(); return RedirectResponse(f"/settings?tab={tab}&message=لا يمكن حذف سجل مرتبط بحركات",303)
    return RedirectResponse(f"/settings?tab={tab}&message=تم حذف البيانات",303)


@app.get("/closing-stock", response_class=HTMLResponse)
def closing_stock(request: Request):
    return render(request, "closing_stock.html", "مخزون آخر الفترة", "closing_stock")


@app.get("/sales", response_class=HTMLResponse)
def sales_entry(
    request: Request,
    edit_id: Optional[int] = None,
    search: str = "",
    branch_id: Optional[str] = None,
    journal_date: Optional[str] = None,
    db: Session = Depends(get_db),
):
    selected_branch_id = int(branch_id) if branch_id and str(branch_id).isdigit() else None
    branches = db.query(Branch).filter(Branch.is_active.is_(True)).order_by(Branch.code).all()
    employees = db.query(Employee).filter(Employee.is_active.is_(True)).order_by(Employee.branch_id, Employee.name).all()
    journal = None
    if edit_id:
        journal = db.query(SalesJournal).options(joinedload(SalesJournal.lines).joinedload(SalesLine.employee),joinedload(SalesJournal.treasury_deposit)).filter(SalesJournal.id == edit_id).first()
        if not journal:
            raise HTTPException(404, "اليومية غير موجودة")

    query = db.query(SalesJournal).options(joinedload(SalesJournal.branch), joinedload(SalesJournal.lines))
    if search:
        query = query.filter(SalesJournal.journal_no.contains(search))
    if selected_branch_id:
        query = query.filter(SalesJournal.branch_id == selected_branch_id)
    if journal_date:
        query = query.filter(SalesJournal.journal_date == parse_date(journal_date))
    results = query.order_by(SalesJournal.journal_date.desc(), SalesJournal.id.desc()).limit(100).all()

    return render(
        request, "sales/index.html", "يومية المبيعات", "sales",
        branches=branches, employees=employees, treasuries=db.query(Treasury).filter(Treasury.is_active.is_(True)).order_by(Treasury.name).all(), journal=journal, results=results,
        is_readonly=bool(journal and journal.status == "posted"),
        today=date.today().strftime("%d/%m/%Y"), selected_branch=selected_branch_id, search=search,
        selected_date=journal_date or "", message=request.query_params.get("message", ""),
    )


@app.post("/sales/save")
async def save_sales(request: Request, db: Session = Depends(get_db)):
    form = await request.form()
    journal_id = int(form.get("journal_id") or 0)
    journal_date = parse_date(str(form.get("journal_date") or ""), date.today())
    branch_id = int(form.get("branch_id") or 0)
    notes = str(form.get("notes") or "").strip()
    treasury_id = int(form.get("treasury_id") or 0)

    if not branch_id or not treasury_id:
        return RedirectResponse("/sales?message=يجب اختيار الفرع والخزينة", status_code=303)
    treasury=db.query(Treasury).filter(Treasury.id==treasury_id,Treasury.is_active.is_(True)).first()
    if not treasury or treasury.branch_id!=branch_id:
        return RedirectResponse("/sales?message=الخزينة المختارة لا تتبع الفرع",status_code=303)

    employee_ids = form.getlist("employee_id")
    shift_values = form.getlist("shift_value")
    discounts = form.getlist("discount")
    net_cash_values = form.getlist("net_cash")
    differences = form.getlist("cash_difference")

    valid_lines = []
    for index, employee_id in enumerate(employee_ids):
        if not employee_id:
            continue
        shift = float(shift_values[index] or 0)
        discount = float(discounts[index] or 0)
        difference = float(differences[index] or 0)
        net_cash = shift - discount + difference
        if shift == 0 and discount == 0 and net_cash == 0 and difference == 0:
            continue
        valid_lines.append((int(employee_id), shift, discount, net_cash, difference))

    if not valid_lines:
        return RedirectResponse("/sales?message=أدخل حركة مبيعات واحدة على الأقل", status_code=303)

    valid_employee_ids = {
        row.id for row in db.query(Employee).filter(
            Employee.id.in_([line[0] for line in valid_lines]),
            Employee.branch_id == branch_id,
            Employee.is_active.is_(True),
        ).all()
    }
    if len(valid_employee_ids) != len({line[0] for line in valid_lines}):
        return RedirectResponse("/sales?message=يوجد مستخدم غير تابع للفرع المختار", status_code=303)

    if journal_id:
        journal = db.query(SalesJournal).options(joinedload(SalesJournal.lines)).filter(SalesJournal.id == journal_id).first()
        if not journal or journal.status == "posted":
            return RedirectResponse("/sales?message=لا يمكن تعديل هذه اليومية", status_code=303)
        journal.journal_date = journal_date
        journal.branch_id = branch_id
        journal.notes = notes
        journal.lines.clear()
    else:
        duplicate = db.query(SalesJournal).filter(SalesJournal.branch_id == branch_id, SalesJournal.journal_date == journal_date).first()
        if duplicate:
            return RedirectResponse(f"/sales?edit_id={duplicate.id}&message=توجد يومية لهذا الفرع والتاريخ وتم فتحها للتعديل", status_code=303)
        journal = SalesJournal(journal_no=next_journal_no(db, journal_date), journal_date=journal_date, branch_id=branch_id, notes=notes)
        db.add(journal)

    for employee_id, shift, discount, net_cash, difference in valid_lines:
        journal.lines.append(SalesLine(employee_id=employee_id, shift_value=shift, discount=discount, net_cash=net_cash, cash_difference=difference))

    total_net=sum(line[3] for line in valid_lines)
    if journal.treasury_deposit:
        journal.treasury_deposit.treasury_id=treasury_id; journal.treasury_deposit.amount=total_net
    else:
        journal.treasury_deposit=TreasuryDeposit(treasury_id=treasury_id,amount=total_net)

    db.commit()
    return RedirectResponse("/sales?message=تم حفظ يومية المبيعات بنجاح", status_code=303)


@app.get("/review", response_class=HTMLResponse)
def unified_review(request: Request, type: str="sales", search: str="", db: Session=Depends(get_db)):
    allowed={"sales","purchases","expenses","other_accounts","supplier_payments"}
    if type not in allowed:type="sales"
    if type=="purchases":
        all_items=db.query(PurchaseJournal).options(joinedload(PurchaseJournal.lines).joinedload(PurchaseLine.supplier)).order_by(PurchaseJournal.journal_date.desc(),PurchaseJournal.id.desc()).all()
    elif type=="expenses":all_items=db.query(ExpenseJournal).options(joinedload(ExpenseJournal.branch),joinedload(ExpenseJournal.lines).joinedload(ExpenseLine.expense_item)).order_by(ExpenseJournal.journal_date.desc(),ExpenseJournal.id.desc()).all()
    elif type=="other_accounts":all_items=db.query(OtherAccountJournal).options(joinedload(OtherAccountJournal.lines).joinedload(OtherAccountLine.account)).order_by(OtherAccountJournal.journal_date.desc(),OtherAccountJournal.id.desc()).all()
    elif type=="sales":all_items=db.query(SalesJournal).options(joinedload(SalesJournal.branch),joinedload(SalesJournal.lines).joinedload(SalesLine.employee)).order_by(SalesJournal.journal_date.desc(),SalesJournal.id.desc()).all()
    else:all_items=[]
    items=[x for x in all_items if x.status=="draft" and (not search or search.lower() in x.journal_no.lower())]
    totals={"draft":len(items),"lines":sum(len(x.lines) for x in items),"value":sum((x.total_net_cash if type=="sales" else x.total_effect if type=="purchases" else x.total_amount) for x in items)}
    return render(request,"sales/review.html","مراجعة وترحيل","review",journals=items,totals=totals,selected_type=type,search=search,message=request.query_params.get("message",""))

@app.get("/sales/review")
def sales_review_redirect():
    return RedirectResponse("/review", status_code=303)


@app.post("/sales/{journal_id}/post")
def post_sales(journal_id: int, db: Session = Depends(get_db)):
    journal = db.query(SalesJournal).filter(SalesJournal.id == journal_id).first()
    if not journal:
        raise HTTPException(404, "اليومية غير موجودة")
    if journal.status != "posted":
        journal.status = "posted"
        journal.posted_at = datetime.utcnow()
        db.commit()
    return RedirectResponse("/review?message=تم ترحيل اليومية بنجاح", status_code=303)


@app.get("/reports",response_class=HTMLResponse)
def unified_reports(request:Request,tab:str="sales",branch_id:Optional[int]=None,employee_id:Optional[int]=None,supplier_id:Optional[int]=None,expense_item_id:Optional[int]=None,account_id:Optional[int]=None,expense_type:str="",entry_type:str="all",date_from:Optional[str]=None,date_to:Optional[str]=None,status:str="all",journal_no:str="",document_no:str="",db:Session=Depends(get_db)):
    allowed={"sales","purchases","expenses","other_accounts","supplier_payments"}; tab=tab if tab in allowed else "sales"
    filters={"branch_id":branch_id,"employee_id":employee_id,"supplier_id":supplier_id,"expense_item_id":expense_item_id,"account_id":account_id,"expense_type":expense_type,"entry_type":entry_type,"date_from":date_from or "","date_to":date_to or "","status":status,"journal_no":journal_no,"document_no":document_no}
    account_summary=None
    if tab=="sales":
        q=db.query(SalesLine).join(SalesJournal).options(joinedload(SalesLine.employee),joinedload(SalesLine.journal).joinedload(SalesJournal.branch))
        if branch_id:q=q.filter(SalesJournal.branch_id==branch_id)
        if employee_id:q=q.filter(SalesLine.employee_id==employee_id)
        if journal_no:q=q.filter(SalesJournal.journal_no.contains(journal_no))
        if date_from:q=q.filter(SalesJournal.journal_date>=parse_date(date_from))
        if date_to:q=q.filter(SalesJournal.journal_date<=parse_date(date_to))
        if status in {"draft","posted"}:q=q.filter(SalesJournal.status==status)
        lines=q.order_by(SalesJournal.journal_date.desc(),SalesLine.id.desc()).all(); totals=[("الورديات",len(lines)),("المبيعات",sum(x.shift_value for x in lines)),("الخصومات",sum(x.discount for x in lines)),("فروق الخزينة",sum(x.cash_difference for x in lines)),("صافي النقدية",sum(x.net_cash for x in lines))]
    elif tab=="purchases":
        q=db.query(PurchaseLine).join(PurchaseJournal).options(joinedload(PurchaseLine.supplier),joinedload(PurchaseLine.journal))
        if supplier_id:q=q.filter(PurchaseLine.supplier_id==supplier_id)
        if entry_type in {"purchase","notice"}:q=q.filter(PurchaseJournal.entry_type==entry_type)
        if journal_no:q=q.filter(PurchaseJournal.journal_no.contains(journal_no))
        if document_no:q=q.filter(PurchaseLine.document_no.contains(document_no))
        if date_from:q=q.filter(PurchaseJournal.journal_date>=parse_date(date_from))
        if date_to:q=q.filter(PurchaseJournal.journal_date<=parse_date(date_to))
        if status in {"draft","posted"}:q=q.filter(PurchaseJournal.status==status)
        lines=q.order_by(PurchaseJournal.journal_date.desc(),PurchaseLine.id.desc()).all(); totals=[("الحركات",len(lines)),("القيمة صيدلي",sum(x.pharmacy_value for x in lines)),("القيمة جمهور",sum(x.public_value for x in lines)),("تأثير المورد",sum(x.account_effect for x in lines))]
    elif tab=="expenses":
        q=db.query(ExpenseLine).join(ExpenseJournal).options(joinedload(ExpenseLine.expense_item),joinedload(ExpenseLine.journal).joinedload(ExpenseJournal.branch))
        if branch_id:q=q.filter(ExpenseJournal.branch_id==branch_id)
        if expense_type in {"operating","general"}:q=q.filter(ExpenseJournal.expense_type==expense_type)
        if expense_item_id:q=q.filter(ExpenseLine.expense_item_id==expense_item_id)
        if journal_no:q=q.filter(ExpenseJournal.journal_no.contains(journal_no))
        if date_from:q=q.filter(ExpenseJournal.journal_date>=parse_date(date_from))
        if date_to:q=q.filter(ExpenseJournal.journal_date<=parse_date(date_to))
        if status in {"draft","posted"}:q=q.filter(ExpenseJournal.status==status)
        lines=q.order_by(ExpenseJournal.journal_date.desc(),ExpenseLine.id.desc()).all();totals=[("الحركات",len(lines)),("إجمالي المصروفات",sum(x.amount for x in lines))]
    elif tab=="other_accounts":
        q=db.query(OtherAccountLine).join(OtherAccountJournal).options(joinedload(OtherAccountLine.account),joinedload(OtherAccountLine.journal))
        if account_id:q=q.filter(OtherAccountLine.account_id==account_id)
        if journal_no:q=q.filter(OtherAccountJournal.journal_no.contains(journal_no))
        if date_from:q=q.filter(OtherAccountJournal.journal_date>=parse_date(date_from))
        if date_to:q=q.filter(OtherAccountJournal.journal_date<=parse_date(date_to))
        if status in {"draft","posted"}:q=q.filter(OtherAccountJournal.status==status)
        lines=q.order_by(OtherAccountJournal.journal_date.desc(),OtherAccountLine.id.desc()).all();totals=[("الحركات",len(lines)),("إجمالي المبالغ",sum(x.amount for x in lines))]
        if account_id:
            account=db.get(OtherAccountItem,account_id)
            if account:
                opening=(account.opening_debit or 0)-(account.opening_credit or 0);funding=sum(x.amount for x in lines if x.journal.transaction_type=="funding");withdrawal=sum(x.amount for x in lines if x.journal.transaction_type=="withdrawal")
                account_summary=[("الرصيد الافتتاحي",opening),("إجمالي التمويل",funding),("إجمالي السحب",withdrawal),("الرصيد الحالي",opening+funding-withdrawal)]
    else:lines=[]; totals=[("الحركات",0),("الإجمالي",0),("المرحل",0),("غير المرحل",0)]
    return render(request,"reports/unified.html","التقارير","reports",tab=tab,lines=lines,totals=totals,account_summary=account_summary,filters=filters,branches=db.query(Branch).order_by(Branch.name).all(),employees=db.query(Employee).order_by(Employee.name).all(),suppliers=db.query(Supplier).order_by(Supplier.name).all(),expense_items=db.query(ExpenseItem).order_by(ExpenseItem.name).all(),accounts=db.query(OtherAccountItem).order_by(OtherAccountItem.name).all())

@app.get("/reports/sales")
def sales_report_redirect():
    return RedirectResponse("/reports?tab=sales",status_code=303)

def report_lines_query(db: Session, branch_id=None, employee_id=None, date_from=None, date_to=None, status="all"):
    query = db.query(SalesLine).join(SalesJournal).options(joinedload(SalesLine.employee), joinedload(SalesLine.journal).joinedload(SalesJournal.branch))
    if branch_id:
        query = query.filter(SalesJournal.branch_id == branch_id)
    if employee_id:
        query = query.filter(SalesLine.employee_id == employee_id)
    if date_from:
        query = query.filter(SalesJournal.journal_date >= parse_date(date_from))
    if date_to:
        query = query.filter(SalesJournal.journal_date <= parse_date(date_to))
    if status in {"draft", "posted"}:
        query = query.filter(SalesJournal.status == status)
    return query.order_by(SalesJournal.journal_date, SalesJournal.id, SalesLine.id).all()


@app.get("/reports/sales/export")
def export_sales_report(
    branch_id: Optional[int] = None, employee_id: Optional[int] = None,
    date_from: Optional[str] = None, date_to: Optional[str] = None,
    status: str = "all", db: Session = Depends(get_db),
):
    lines = report_lines_query(db, branch_id, employee_id, date_from, date_to, status)
    wb = Workbook()
    ws = wb.active
    ws.title = "تقرير المبيعات"
    ws.sheet_view.rightToLeft = True
    headers = ["رقم اليومية", "التاريخ", "الفرع", "كود الموظف", "المستخدم", "قيمة الوردية", "الخصم", "صافي النقدية", "فروق الخزينة", "الحالة"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="243B53")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for line in lines:
        ws.append([
            line.journal.journal_no, line.journal.journal_date.strftime("%d/%m/%Y"), line.journal.branch.name,
            line.employee.code, line.employee.name, line.shift_value, line.discount, line.net_cash,
            line.cash_difference, "مرحلة" if line.journal.status == "posted" else "غير مرحلة",
        ])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    for column, width in enumerate([18, 14, 22, 15, 28, 16, 14, 16, 16, 14], start=1):
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.freeze_panes = "A2"
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    headers_out = {"Content-Disposition": "attachment; filename=AlFarouq_Sales_Report.xlsx"}
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers_out)



def next_purchase_no(db: Session, d: date) -> str:
    prefix=f"PJ-{d.year}-"; last=db.query(func.max(PurchaseJournal.journal_no)).filter(PurchaseJournal.journal_no.like(f"{prefix}%")).scalar(); seq=int(last.rsplit("-",1)[-1])+1 if last else 1; return f"{prefix}{seq:05d}"

@app.get("/purchases", response_class=HTMLResponse)
def purchases(request: Request, edit_id: Optional[int]=None, search: str="", status: str="all", journal_date:Optional[str]=None, db: Session=Depends(get_db)):
    journal=db.query(PurchaseJournal).options(joinedload(PurchaseJournal.lines).joinedload(PurchaseLine.supplier)).filter(PurchaseJournal.id==edit_id).first() if edit_id else None
    q=db.query(PurchaseJournal).options(joinedload(PurchaseJournal.lines).joinedload(PurchaseLine.supplier))
    if search: q=q.join(PurchaseLine).join(Supplier).filter((PurchaseJournal.journal_no.contains(search)) | (PurchaseLine.document_no.contains(search)) | (Supplier.name.contains(search)))
    if journal_date:q=q.filter(PurchaseJournal.journal_date==parse_date(journal_date))
    if status in {"draft","posted"}: q=q.filter(PurchaseJournal.status==status)
    results=q.order_by(PurchaseJournal.journal_date.desc(),PurchaseJournal.id.desc()).limit(100).all()
    return render(request,"purchases/index.html","يومية المشتريات","purchases",suppliers=db.query(Supplier).filter(Supplier.is_active.is_(True)).order_by(Supplier.name).all(),notice_types=db.query(OtherAccountItem).filter(OtherAccountItem.is_active.is_(True)).order_by(OtherAccountItem.name).all(),journal=journal,results=results,today=date.today().strftime("%d/%m/%Y"),readonly=bool(journal and journal.status=="posted"),search=search,selected_status=status,selected_date=journal_date or "",message=request.query_params.get("message",""))

@app.post("/purchases/save")
async def save_purchases(request: Request, db: Session=Depends(get_db)):
    f=await request.form(); jid=int(f.get("journal_id") or 0); d=parse_date(str(f.get("journal_date") or ""),date.today()); et=str(f.get("entry_type") or "purchase")
    fixed_notice_types={"مرتجع":-1,"لم يصل":-1,"خصم إضافي":-1,"ت. إضافية":1,"غرامة":-1,"أخرى":1}
    journal_notice_type=str(f.get("notice_type") or "").strip() if et=="notice" else ""
    if et not in {"purchase","notice"}: return RedirectResponse("/purchases?message=اختر نوع اليومية",303)
    if et=="notice" and journal_notice_type not in fixed_notice_types: return RedirectResponse("/purchases?message=اختر نوع الإشعار",303)
    supplier_ids=f.getlist("supplier_id"); docs=f.getlist("document_no"); pvs=f.getlist("pharmacy_value"); pubs=f.getlist("public_value"); descs=f.getlist("description"); notes=f.getlist("line_notes")
    rows=[]
    for i,sid in enumerate(supplier_ids):
        if not sid or not docs[i].strip(): continue
        pv=float(pvs[i] or 0); pub=float(pubs[i] or 0); disc=abs(((pv/pub)*100)-100) if pub else 0; nt=""; effect=pv
        if et=="notice":
            nt=journal_notice_type
            if pv <= 0:return RedirectResponse("/purchases?message=أدخل القيمة صيدلي للإشعار",303)
            if pub <= 0:return RedirectResponse("/purchases?message=أدخل القيمة جمهور للإشعار",303)
            if not (descs[i].strip() if i<len(descs) else ""):return RedirectResponse("/purchases?message=أدخل البيان لكل سطر إشعار",303)
            effect=abs(pv) * fixed_notice_types[nt]
        rows.append((int(sid),docs[i].strip(),nt,pv,pub,max(disc,0),effect,descs[i] if i<len(descs) else "",notes[i] if i<len(notes) else ""))
    if not rows: return RedirectResponse("/purchases?message=أدخل حركة واحدة على الأقل",303)
    for sid,doc,*_ in rows:
        dup=db.query(PurchaseLine).filter(PurchaseLine.supplier_id==sid,PurchaseLine.document_no==doc,PurchaseLine.entry_type==et)
        if jid: dup=dup.filter(PurchaseLine.journal_id!=jid)
        if dup.first(): return RedirectResponse(f"/purchases?message=رقم المستند {doc} مكرر لنفس المورد",303)
    if jid:
        j=db.query(PurchaseJournal).options(joinedload(PurchaseJournal.lines)).filter(PurchaseJournal.id==jid).first()
        if not j or j.status=="posted": return RedirectResponse("/purchases?message=لا يمكن تعديل اليومية",303)
        j.journal_date=d; j.entry_type=et; j.notice_type=journal_notice_type; j.lines.clear()
    else:
        j=PurchaseJournal(journal_no=next_purchase_no(db,d),journal_date=d,entry_type=et,notice_type=journal_notice_type); db.add(j)
    for sid,doc,nt,pv,pub,disc,effect,des,ntes in rows: j.lines.append(PurchaseLine(supplier_id=sid,entry_type=et,document_no=doc,notice_type=nt,pharmacy_value=pv,public_value=pub,discount_percent=disc,account_effect=effect,description=des,notes=ntes))
    db.commit(); return RedirectResponse("/purchases?message=تم حفظ اليومية بنجاح",303)

@app.post("/purchases/{journal_id}/post")
def post_purchase(journal_id:int,db:Session=Depends(get_db)):
    j=db.query(PurchaseJournal).filter(PurchaseJournal.id==journal_id).first();
    if not j: raise HTTPException(404,"اليومية غير موجودة")
    j.status="posted"; j.posted_at=datetime.utcnow(); db.commit(); return RedirectResponse("/review?type=purchases&message=تم الترحيل",303)

@app.get("/reports/purchases")
def purchase_report_redirect():
    return RedirectResponse("/reports?tab=purchases",status_code=303)

@app.get("/reports/purchases/export")
def export_purchases(supplier_id:Optional[int]=None,entry_type:str="all",status:str="all",document_no:str="",date_from:Optional[str]=None,date_to:Optional[str]=None,db:Session=Depends(get_db)):
    q=db.query(PurchaseLine).join(PurchaseJournal).options(joinedload(PurchaseLine.supplier),joinedload(PurchaseLine.journal));
    if supplier_id:q=q.filter(PurchaseLine.supplier_id==supplier_id)
    if entry_type in {"purchase","notice"}:q=q.filter(PurchaseLine.entry_type==entry_type)
    if status in {"draft","posted"}:q=q.filter(PurchaseJournal.status==status)
    if document_no:q=q.filter(PurchaseLine.document_no.contains(document_no))
    if date_from:q=q.filter(PurchaseJournal.journal_date>=parse_date(date_from))
    if date_to:q=q.filter(PurchaseJournal.journal_date<=parse_date(date_to))
    lines=q.order_by(PurchaseJournal.journal_date,PurchaseLine.id).all(); wb=Workbook(); ws=wb.active; ws.title="تقرير المشتريات"; ws.sheet_view.rightToLeft=True; headers=["اليومية","التاريخ","النوع","المورد","رقم المستند","نوع الإشعار","صيدلي","جمهور","الخصم %","تأثير الحساب","الوصف","الحالة"]; ws.append(headers)
    for c in ws[1]:c.font=Font(bold=True,color="FFFFFF");c.fill=PatternFill("solid",fgColor="7C3AED");c.alignment=Alignment(horizontal="center")
    for x in lines:ws.append([x.journal.journal_no,x.journal.journal_date.strftime("%d/%m/%Y"),"فاتورة شراء" if x.entry_type=="purchase" else (x.journal.notice_type or x.notice_type or "إشعار"),x.supplier.name,x.document_no,x.notice_type,x.pharmacy_value,x.public_value,x.discount_percent,x.account_effect,x.description,"مرحلة" if x.journal.status=="posted" else "غير مرحلة"] )
    for i,w in enumerate([18,14,12,24,18,18,14,14,12,16,28,14],1):ws.column_dimensions[get_column_letter(i)].width=w
    out=BytesIO();wb.save(out);out.seek(0);return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":"attachment; filename=AlFarouq_Purchases_Report.xlsx"})

def next_number(db:Session, model, prefix:str, d:date)->str:
    stem=f"{prefix}-{d.year}-"; last=db.query(func.max(model.journal_no)).filter(model.journal_no.like(f"{stem}%")).scalar(); seq=int(last.rsplit("-",1)[-1])+1 if last else 1; return f"{stem}{seq:05d}"

@app.get("/expenses",response_class=HTMLResponse)
def expenses(request:Request,edit_id:Optional[int]=None,search:str="",status:str="all",journal_date:Optional[str]=None,db:Session=Depends(get_db)):
    journal=db.query(ExpenseJournal).options(joinedload(ExpenseJournal.branch),joinedload(ExpenseJournal.lines).joinedload(ExpenseLine.expense_item),joinedload(ExpenseJournal.treasury_payment)).filter(ExpenseJournal.id==edit_id).first() if edit_id else None
    if edit_id and not journal:raise HTTPException(404,"اليومية غير موجودة")
    q=db.query(ExpenseJournal).options(joinedload(ExpenseJournal.branch),joinedload(ExpenseJournal.lines))
    if search:q=q.filter(ExpenseJournal.journal_no.contains(search))
    if status in {"draft","posted"}:q=q.filter(ExpenseJournal.status==status)
    if journal_date:q=q.filter(ExpenseJournal.journal_date==parse_date(journal_date))
    return render(request,"expenses/index.html","يومية مصروفات الفروع","expenses",journal=journal,results=q.order_by(ExpenseJournal.journal_date.desc(),ExpenseJournal.id.desc()).limit(100).all(),branches=db.query(Branch).filter(Branch.is_active.is_(True)).order_by(Branch.name).all(),treasuries=db.query(Treasury).filter(Treasury.is_active.is_(True)).order_by(Treasury.name).all(),expense_items=db.query(ExpenseItem).filter(ExpenseItem.is_active.is_(True)).order_by(ExpenseItem.name).all(),readonly=bool(journal and journal.status=="posted"),choose_treasury=request.query_params.get("choose_treasury")=="1",today=date.today().strftime("%d/%m/%Y"),search=search,selected_status=status,selected_date=journal_date or "",message=request.query_params.get("message",""))

@app.post("/expenses/save")
async def save_expenses(request:Request,db:Session=Depends(get_db)):
    f=await request.form(); jid=int(f.get("journal_id") or 0); d=parse_date(f.get("journal_date"),date.today()); kind=str(f.get("expense_type") or ""); branch_id=int(f.get("branch_id") or 0) or None
    if kind not in {"operating","general"}:return RedirectResponse("/expenses?message=اختر نوع المصروف",303)
    if kind=="operating" and not branch_id:return RedirectResponse("/expenses?message=الفرع مطلوب للمصروف التشغيلي",303)
    if kind=="general":branch_id=None
    ids=f.getlist("expense_item_id"); amounts=f.getlist("amount"); notes=f.getlist("line_notes"); rows=[]
    for i,item_id in enumerate(ids):
        amount=float(amounts[i] or 0)
        if item_id and amount>0:rows.append((int(item_id),amount,notes[i].strip() if i<len(notes) else ""))
    if not rows:return RedirectResponse("/expenses?message=أدخل حركة مصروف واحدة على الأقل",303)
    valid={x.id for x in db.query(ExpenseItem).filter(ExpenseItem.id.in_([x[0] for x in rows]),ExpenseItem.expense_type==kind,ExpenseItem.is_active.is_(True)).all()}
    if len(valid)!=len(set(x[0] for x in rows)):return RedirectResponse("/expenses?message=يوجد حساب لا يطابق نوع المصروف",303)
    if jid:
        j=db.query(ExpenseJournal).options(joinedload(ExpenseJournal.lines)).filter(ExpenseJournal.id==jid).first()
        if not j or j.status=="posted":return RedirectResponse("/expenses?message=لا يمكن تعديل اليومية",303)
        j.journal_date=d;j.expense_type=kind;j.branch_id=branch_id;j.lines.clear()
    else:j=ExpenseJournal(journal_no=next_number(db,ExpenseJournal,"EJ",d),journal_date=d,expense_type=kind,branch_id=branch_id);db.add(j)
    for item_id,amount,note in rows:j.lines.append(ExpenseLine(expense_item_id=item_id,amount=amount,notes=note))
    db.commit();db.refresh(j);return RedirectResponse(f"/expenses?edit_id={j.id}&choose_treasury=1&message=تم حفظ اليومية، اختر خزينة الصرف",303)

@app.post("/expenses/{journal_id}/treasury")
async def save_expense_treasury(journal_id:int,request:Request,db:Session=Depends(get_db)):
    f=await request.form();treasury_id=int(f.get("treasury_id") or 0)
    j=db.query(ExpenseJournal).options(joinedload(ExpenseJournal.treasury_payment)).filter(ExpenseJournal.id==journal_id).first();treasury=db.query(Treasury).filter(Treasury.id==treasury_id,Treasury.is_active.is_(True)).first()
    if not j or j.status=="posted":return RedirectResponse("/expenses?message=لا يمكن تعديل حركة خزينة هذه اليومية",303)
    if not treasury:return RedirectResponse(f"/expenses?edit_id={journal_id}&choose_treasury=1&message=اختر خزينة صحيحة",303)
    if j.expense_type=="operating" and treasury.branch_id!=j.branch_id:return RedirectResponse(f"/expenses?edit_id={journal_id}&choose_treasury=1&message=الخزينة المختارة لا تتبع الفرع",303)
    if j.treasury_payment:j.treasury_payment.treasury_id=treasury.id;j.treasury_payment.amount=j.total_amount
    else:j.treasury_payment=ExpenseTreasuryPayment(treasury_id=treasury.id,amount=j.total_amount)
    db.commit();return RedirectResponse(f"/expenses?edit_id={journal_id}&message=تم إنشاء حركة الصرف من الخزينة بنجاح",303)

@app.post("/expenses/{journal_id}/post")
def post_expense(journal_id:int,db:Session=Depends(get_db)):
    j=db.get(ExpenseJournal,journal_id)
    if not j:raise HTTPException(404,"اليومية غير موجودة")
    if j.status!="posted":j.status="posted";j.posted_at=datetime.utcnow();db.commit()
    return RedirectResponse("/review?type=expenses&message=تم الترحيل",303)

@app.get("/other-accounts",response_class=HTMLResponse)
def other_accounts(request:Request,edit_id:Optional[int]=None,search:str="",status:str="all",journal_date:Optional[str]=None,db:Session=Depends(get_db)):
    journal=db.query(OtherAccountJournal).options(joinedload(OtherAccountJournal.lines).joinedload(OtherAccountLine.account)).filter(OtherAccountJournal.id==edit_id).first() if edit_id else None
    if edit_id and not journal:raise HTTPException(404,"اليومية غير موجودة")
    q=db.query(OtherAccountJournal).options(joinedload(OtherAccountJournal.lines))
    if search:q=q.filter(OtherAccountJournal.journal_no.contains(search))
    if status in {"draft","posted"}:q=q.filter(OtherAccountJournal.status==status)
    if journal_date:q=q.filter(OtherAccountJournal.journal_date==parse_date(journal_date))
    return render(request,"other_accounts/index.html","يومية الحسابات الأخرى","other_accounts",journal=journal,results=q.order_by(OtherAccountJournal.journal_date.desc(),OtherAccountJournal.id.desc()).limit(100).all(),accounts=db.query(OtherAccountItem).filter(OtherAccountItem.is_active.is_(True)).order_by(OtherAccountItem.name).all(),readonly=bool(journal and journal.status=="posted"),today=date.today().strftime("%d/%m/%Y"),search=search,selected_status=status,selected_date=journal_date or "",message=request.query_params.get("message",""))

@app.post("/other-accounts/save")
async def save_other_accounts(request:Request,db:Session=Depends(get_db)):
    f=await request.form();jid=int(f.get("journal_id") or 0);d=parse_date(f.get("journal_date"),date.today());kind=str(f.get("transaction_type") or "")
    if kind not in {"funding","withdrawal"}:return RedirectResponse("/other-accounts?message=اختر نوع الحركة",303)
    ids=f.getlist("account_id");amounts=f.getlist("amount");descriptions=f.getlist("description");rows=[]
    for i,account_id in enumerate(ids):
        amount=float(amounts[i] or 0)
        if account_id and amount>0:rows.append((int(account_id),amount,descriptions[i].strip() if i<len(descriptions) else ""))
    if not rows:return RedirectResponse("/other-accounts?message=أدخل حركة واحدة على الأقل",303)
    valid={x.id for x in db.query(OtherAccountItem).filter(OtherAccountItem.id.in_([x[0] for x in rows]),OtherAccountItem.is_active.is_(True)).all()}
    if len(valid)!=len(set(x[0] for x in rows)):return RedirectResponse("/other-accounts?message=يوجد حساب غير صالح",303)
    if jid:
        j=db.query(OtherAccountJournal).options(joinedload(OtherAccountJournal.lines)).filter(OtherAccountJournal.id==jid).first()
        if not j or j.status=="posted":return RedirectResponse("/other-accounts?message=لا يمكن تعديل اليومية",303)
        j.journal_date=d;j.transaction_type=kind;j.lines.clear()
    else:j=OtherAccountJournal(journal_no=next_number(db,OtherAccountJournal,"OJ",d),journal_date=d,transaction_type=kind);db.add(j)
    for account_id,amount,description in rows:j.lines.append(OtherAccountLine(account_id=account_id,amount=amount,description=description))
    db.commit();return RedirectResponse("/other-accounts?message=تم حفظ يومية الحسابات الأخرى بنجاح",303)

@app.post("/other-accounts/{journal_id}/post")
def post_other_account(journal_id:int,db:Session=Depends(get_db)):
    j=db.get(OtherAccountJournal,journal_id)
    if not j:raise HTTPException(404,"اليومية غير موجودة")
    if j.status!="posted":j.status="posted";j.posted_at=datetime.utcnow();db.commit()
    return RedirectResponse("/review?type=other_accounts&message=تم الترحيل",303)

def journal_export(title:str,headers:list,rows:list,filename:str):
    wb=Workbook();ws=wb.active;ws.title=title;ws.sheet_view.rightToLeft=True;ws.append(headers)
    for c in ws[1]:c.font=Font(bold=True,color="FFFFFF");c.fill=PatternFill("solid",fgColor="243B53");c.alignment=Alignment(horizontal="center")
    for row in rows:ws.append(row)
    for column in range(1,len(headers)+1):ws.column_dimensions[get_column_letter(column)].width=22
    out=BytesIO();wb.save(out);out.seek(0);return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":f"attachment; filename={filename}"})

@app.get("/reports/expenses/export")
def export_expenses(branch_id:Optional[int]=None,expense_type:str="",expense_item_id:Optional[int]=None,status:str="all",journal_no:str="",date_from:Optional[str]=None,date_to:Optional[str]=None,db:Session=Depends(get_db)):
    q=db.query(ExpenseLine).join(ExpenseJournal).options(joinedload(ExpenseLine.expense_item),joinedload(ExpenseLine.journal).joinedload(ExpenseJournal.branch))
    if branch_id:q=q.filter(ExpenseJournal.branch_id==branch_id)
    if expense_type in {"operating","general"}:q=q.filter(ExpenseJournal.expense_type==expense_type)
    if expense_item_id:q=q.filter(ExpenseLine.expense_item_id==expense_item_id)
    if status in {"draft","posted"}:q=q.filter(ExpenseJournal.status==status)
    if journal_no:q=q.filter(ExpenseJournal.journal_no.contains(journal_no))
    if date_from:q=q.filter(ExpenseJournal.journal_date>=parse_date(date_from))
    if date_to:q=q.filter(ExpenseJournal.journal_date<=parse_date(date_to))
    rows=[[x.journal.journal_no,x.journal.journal_date.strftime("%d/%m/%Y"),"مصروف تشغيلي" if x.journal.expense_type=="operating" else "مصروف عمومي",x.journal.branch.name if x.journal.branch else "—",x.expense_item.name,x.amount,x.notes,"مرحلة" if x.journal.status=="posted" else "غير مرحلة"] for x in q.order_by(ExpenseJournal.journal_date,ExpenseLine.id).all()]
    return journal_export("تقرير المصروفات",["اليومية","التاريخ","النوع","الفرع","الحساب","القيمة","ملاحظات","الحالة"],rows,"AlFarouq_Expenses_Report.xlsx")

@app.get("/reports/other-accounts/export")
def export_other_accounts(account_id:Optional[int]=None,status:str="all",journal_no:str="",date_from:Optional[str]=None,date_to:Optional[str]=None,db:Session=Depends(get_db)):
    q=db.query(OtherAccountLine).join(OtherAccountJournal).options(joinedload(OtherAccountLine.account),joinedload(OtherAccountLine.journal))
    if account_id:q=q.filter(OtherAccountLine.account_id==account_id)
    if status in {"draft","posted"}:q=q.filter(OtherAccountJournal.status==status)
    if journal_no:q=q.filter(OtherAccountJournal.journal_no.contains(journal_no))
    if date_from:q=q.filter(OtherAccountJournal.journal_date>=parse_date(date_from))
    if date_to:q=q.filter(OtherAccountJournal.journal_date<=parse_date(date_to))
    rows=[[x.journal.journal_no,x.journal.journal_date.strftime("%d/%m/%Y"),"تمويل" if x.journal.transaction_type=="funding" else "سحب أول",x.account.name,x.amount,x.description,"مرحلة" if x.journal.status=="posted" else "غير مرحلة"] for x in q.order_by(OtherAccountJournal.journal_date,OtherAccountLine.id).all()]
    return journal_export("تقرير الحسابات الأخرى",["اليومية","التاريخ","نوع الحركة","الحساب","المبلغ","البيان","الحالة"],rows,"AlFarouq_Other_Accounts_Report.xlsx")

@app.get("/health")
def health():
    return {"status": "ok", "app": "صيدليات الفاروق"}
