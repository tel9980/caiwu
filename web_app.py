#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小型氧化加工厂财务管理系统 - Web版
==================================
基于 Flask 的现代化 Web 界面

依赖安装：
    pip install flask openpyxl

启动：
    python web_app.py
    或双击 启动WEB系统.bat
"""
import os, sys
from datetime import datetime, date
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 氧化加工厂财务系统 import (
    DataManager, EXPENSE_CATEGORIES, MATERIALS, SAMPLE_CUSTOMERS,
    VERSION, fmt, calc_income_tax, EXCEL_FILE, DATA_DIR
)

app = Flask(__name__)
app.secret_key = 'caiwu_system_2026'
app.jinja_env.globals.update(today=lambda: date.today())

def get_dm():
    dm = DataManager()
    return dm

def get_income_data(dm):
    ws = dm.sheet("收入记录")
    rows = dm.read_rows(ws, start_row=4, end_col=10)
    data = []
    for r in rows:
        if r[0] and r[2]:
            data.append({
                "date": r[0], "customer": str(r[2] or ""),
                "amount": to_num(r[3]),
                "method": str(r[4] or ""), "tax_type": str(r[5] or ""),
                "invoice": str(r[9] or ""), "note": ""
            })
    return data

def get_expense_data(dm):
    ws = dm.sheet("支出记录")
    rows = dm.read_rows(ws, start_row=4, end_col=7)
    data = []
    for r in rows:
        if r[0] and r[3]:
            data.append({
                "date": r[0], "category": str(r[2] or ""),
                "amount": to_num(r[3]),
                "supplier": str(r[4] or ""), "method": str(r[5] or ""),
                "note": str(r[6] or "")
            })
    return data

def get_arap_data(dm):
    ws = dm.sheet("应收应付")
    rows = dm.read_rows(ws, start_row=4, end_col=7)
    data = []
    for r in rows:
        if r[0]:
            balance = to_num(r[5])  # 期末应收 = 余额
            data.append({
                "customer": str(r[0] or ""), "type": "应收",
                "amount": balance, "paid": to_num(r[3]),
                "balance": balance, "note": str(r[6] or "")
            })
    return data

def get_salary_data(dm):
    ws = dm.sheet("工资表")
    rows = dm.read_rows(ws, start_row=5, end_col=10)
    data = []
    for r in rows:
        if r[1]:
            data.append({
                "name": str(r[1]), "base": to_num(r[2]),
                "overtime": to_num(r[3]), "bonus": to_num(r[4]),
                "social": to_num(r[6]), "tax": to_num(r[8]),
                "net": to_num(r[9])
            })
    return data

def get_material_data(dm):
    ws = dm.sheet("材料进销存台账")
    rows = dm.read_rows(ws, start_row=4, end_col=10)
    data = []
    for r in rows:
        if r[1]:
            data.append({
                "date": r[0], "name": str(r[1] or ""), "spec": str(r[2] or ""),
                "qty": to_num(r[3]), "price": to_num(r[4]),
                "amount": to_num(r[5])
            })
    return data

def to_num(v):
    if v is None: return 0
    if isinstance(v, (int, float)): return v
    try: return float(str(v).replace(',', '').replace('¥', '').replace(' ', ''))
    except: return 0

def calc_summary(dm):
    incomes = get_income_data(dm)
    expenses = get_expense_data(dm)
    arap = get_arap_data(dm)
    salaries = get_salary_data(dm)
    materials = get_material_data(dm)

    total_inc = sum(to_num(i["amount"]) for i in incomes)
    total_exp = sum(to_num(e["amount"]) for e in expenses)
    profit = total_inc - total_exp
    profit_rate = (profit / total_inc * 100) if total_inc else 0
    total_ar = sum(to_num(a["balance"]) for a in arap if a["type"] == "应收")
    total_ap = sum(to_num(a["balance"]) for a in arap if a["type"] == "应付")

    return {
        "inc_count": len(incomes), "exp_count": len(expenses),
        "ar_count": len(arap), "sal_count": len(salaries),
        "mat_count": len(materials),
        "total_inc": total_inc, "total_exp": total_exp,
        "profit": profit, "profit_rate": profit_rate,
        "total_ar": total_ar, "total_ap": total_ap
    }

@app.route('/')
def dashboard():
    dm = get_dm()
    s = calc_summary(dm)
    return render_template('dashboard.html', summary=s, version=VERSION)

@app.route('/income')
def income():
    dm = get_dm()
    data = get_income_data(dm)
    return render_template('income.html', data=data, version=VERSION)

@app.route('/income/add', methods=['POST'])
def income_add():
    dm = get_dm()
    ws = dm.sheet("收入记录")
    row = dm.next_row(ws)
    dm.write_row(ws, row, [
        request.form.get('date', str(date.today())),
        request.form.get('customer', ''),
        float(request.form.get('amount', 0)),
        request.form.get('method', '银行转账'),
        request.form.get('tax_type', '含税'),
        request.form.get('invoice', '否'),
        request.form.get('note', '')
    ])
    dm.save()
    flash('收入已添加', 'success')
    return redirect(url_for('income'))

@app.route('/expense')
def expense():
    dm = get_dm()
    data = get_expense_data(dm)
    return render_template('expense.html', data=data, categories=EXPENSE_CATEGORIES, version=VERSION)

@app.route('/expense/add', methods=['POST'])
def expense_add():
    dm = get_dm()
    ws = dm.sheet("支出记录")
    row = dm.next_row(ws)
    dm.write_row(ws, row, [
        request.form.get('date', str(date.today())),
        request.form.get('category', ''),
        float(request.form.get('amount', 0)),
        request.form.get('supplier', ''),
        request.form.get('method', '银行转账'),
        request.form.get('note', '')
    ])
    dm.save()
    flash('支出已添加', 'success')
    return redirect(url_for('expense'))

@app.route('/arap')
def arap():
    dm = get_dm()
    data = get_arap_data(dm)
    return render_template('arap.html', data=data, version=VERSION)

@app.route('/report/profit')
def profit_report():
    dm = get_dm()
    s = calc_summary(dm)
    expenses = get_expense_data(dm)
    by_category = {}
    for e in expenses:
        cat = e["category"] or "其他"
        by_category[cat] = by_category.get(cat, 0) + (e["amount"] or 0)
    cat_list = sorted(by_category.items(), key=lambda x: -x[1])
    return render_template('profit.html', summary=s, categories=cat_list, version=VERSION)

@app.route('/report/balance')
def balance_report():
    dm = get_dm()
    s = calc_summary(dm)
    return render_template('balance.html', summary=s, version=VERSION)

@app.route('/report/monthly')
def monthly_report():
    dm = get_dm()
    incomes = get_income_data(dm)
    expenses = get_expense_data(dm)
    monthly = {}
    for i in incomes:
        d = i["date"]
        if d and hasattr(d, 'month'):
            key = f"{d.year}-{d.month:02d}"
            monthly.setdefault(key, {"收入": 0, "支出": 0})
            monthly[key]["收入"] += i["amount"] or 0
    for e in expenses:
        d = e["date"]
        if d and hasattr(d, 'month'):
            key = f"{d.year}-{d.month:02d}"
            monthly.setdefault(key, {"收入": 0, "支出": 0})
            monthly[key]["支出"] += e["amount"] or 0

    months = sorted(monthly.keys())
    chart_labels = months
    chart_income = [monthly[m]["收入"] for m in months]
    chart_expense = [monthly[m]["支出"] for m in months]
    chart_profit = [monthly[m]["收入"] - monthly[m]["支出"] for m in months]

    return render_template('monthly.html',
        months=months, monthly=monthly,
        labels=chart_labels, income_data=chart_income,
        expense_data=chart_expense, profit_data=chart_profit,
        version=VERSION
    )

@app.route('/salary')
def salary():
    dm = get_dm()
    data = get_salary_data(dm)
    return render_template('salary.html', data=data, version=VERSION)

@app.route('/material')
def material():
    dm = get_dm()
    data = get_material_data(dm)
    return render_template('material.html', data=data, materials=MATERIALS, version=VERSION)

@app.route('/api/summary')
def api_summary():
    dm = get_dm()
    s = calc_summary(dm)
    return jsonify(s)

@app.route('/api/income')
def api_income():
    dm = get_dm()
    return jsonify(get_income_data(dm))

@app.route('/api/income/search')
def api_income_search():
    dm = get_dm()
    q = request.args.get('q', '').lower()
    data = get_income_data(dm)
    if q:
        data = [d for d in data if q in str(d.get('customer','')).lower() or q in str(d.get('date',''))]
    return jsonify(data)

@app.route('/api/expense/search')
def api_expense_search():
    dm = get_dm()
    q = request.args.get('q', '').lower()
    cat = request.args.get('cat', '')
    data = get_expense_data(dm)
    if q:
        data = [d for d in data if q in str(d.get('category','')).lower() or q in str(d.get('supplier','')).lower()]
    if cat:
        data = [d for d in data if d.get('category') == cat]
    return jsonify(data)

@app.route('/api/income/delete', methods=['POST'])
def api_income_delete():
    dm = get_dm()
    idx = request.json.get('index', -1)
    ws = dm.sheet("收入记录")
    row = idx + 4
    if 4 <= row <= dm.last_row(ws):
        ws.delete_rows(row)
        dm.save()
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "无效行号"}), 400

@app.route('/api/expense/delete', methods=['POST'])
def api_expense_delete():
    dm = get_dm()
    idx = request.json.get('index', -1)
    ws = dm.sheet("支出记录")
    row = idx + 4
    if 4 <= row <= dm.last_row(ws):
        ws.delete_rows(row)
        dm.save()
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "无效行号"}), 400

@app.route('/api/customers')
def api_customers():
    dm = get_dm()
    ws = dm.sheet("收入记录")
    seen = set()
    customers = []
    for r in dm.read_rows(ws, start_row=4, end_col=3):
        if r[2] and str(r[2]) not in seen:
            seen.add(str(r[2]))
            customers.append(str(r[2]))
    return jsonify(sorted(customers))

@app.route('/api/alerts')
def api_alerts():
    dm = get_dm()
    alerts = []
    arap = get_arap_data(dm)
    for a in arap:
        if a["type"] == "应收" and a["balance"] > 0:
            if a["balance"] > 20000:
                alerts.append({"type": "danger", "icon": "bi-exclamation-triangle", "msg": f"应收 {a['customer']}：¥{a['balance']:,.2f}（大额应收）"})
            elif a["balance"] > 10000:
                alerts.append({"type": "warning", "icon": "bi-exclamation-circle", "msg": f"应收 {a['customer']}：¥{a['balance']:,.2f}"})

    ws = dm.sheet("材料进销存台账")
    stock = {}
    for r in dm.read_rows(ws, start_row=4, end_col=10):
        name = str(r[1] or "")
        in_qty = to_num(r[3]); out_qty = to_num(r[6])
        bal = stock.get(name, 0) + in_qty - out_qty
        stock[name] = bal
    for name, bal in sorted(stock.items()):
        if name and bal < 10:
            alerts.append({"type": "warning" if bal > 0 else "danger", "icon": "bi-box", "msg": f"材料 {name}：库存 {bal}（{'缺货' if bal <= 0 else '低于安全线'}）"})

    return jsonify(alerts[:8])

@app.route('/export/profit')
def export_profit():
    dm = get_dm()
    s = calc_summary(dm)
    expenses = get_expense_data(dm)
    by_cat = {}
    for e in expenses:
        c = e["category"] or "其他"
        by_cat[c] = by_cat.get(c, 0) + (e["amount"] or 0)

    from openpyxl import Workbook
    from io import BytesIO
    wb = Workbook()
    ws = wb.active
    ws.title = "利润报表"
    ws['A1'] = '利润报表'; ws.merge_cells('A1:C1')
    ws['A3'] = '项目'; ws['B3'] = '金额'
    ws['A4'] = '总收入'; ws['B4'] = s['total_inc']
    ws['A5'] = '总支出'; ws['B5'] = s['total_exp']
    ws['A6'] = '净利润'; ws['B6'] = s['profit']
    ws['A8'] = '支出类别'; ws['B8'] = '金额'
    row = 9
    for cat, amt in sorted(by_cat.items(), key=lambda x: -x[1]):
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2, value=amt)
        row += 1

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read(), 200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="利润报表.xlsx"'
    }

@app.route('/export/monthly')
def export_monthly():
    dm = get_dm()
    incomes = get_income_data(dm)
    expenses = get_expense_data(dm)
    monthly = {}
    for i in incomes:
        d = i["date"]
        if d:
            ds = str(d) if not hasattr(d, 'strftime') else d.strftime('%Y-%m')
            key = ds[:7]
            monthly.setdefault(key, {"收入": 0, "支出": 0})
            monthly[key]["收入"] += i["amount"] or 0
    for e in expenses:
        d = e["date"]
        if d:
            ds = str(d) if not hasattr(d, 'strftime') else d.strftime('%Y-%m')
            key = ds[:7]
            monthly.setdefault(key, {"收入": 0, "支出": 0})
            monthly[key]["支出"] += e["amount"] or 0

    from openpyxl import Workbook
    from io import BytesIO
    wb = Workbook()
    ws = wb.active
    ws.title = "月度报表"
    ws['A1'] = '月度报表'; ws.merge_cells('A1:D1')
    ws['A3'] = '月份'; ws['B3'] = '收入'; ws['C3'] = '支出'; ws['D3'] = '利润'
    row = 4
    for m in sorted(monthly.keys()):
        r = monthly[m]
        ws.cell(row=row, column=1, value=m)
        ws.cell(row=row, column=2, value=r["收入"])
        ws.cell(row=row, column=3, value=r["支出"])
        ws.cell(row=row, column=4, value=r["收入"] - r["支出"])
        row += 1

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read(), 200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="月度报表.xlsx"'
    }

@app.route('/_ah/health')
def health():
    return 'ok'

if __name__ == '__main__':
    print(f"  Web版 {VERSION}")
    print(f"  数据文件: {EXCEL_FILE}")
    print(f"  访问地址: http://127.0.0.1:5000")
    print(f"  按 Ctrl+C 停止服务器")
    app.run(debug=True, host='0.0.0.0', port=5000)
