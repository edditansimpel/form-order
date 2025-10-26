import os
from flask import Flask, render_template, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import pandas as pd
import io

app = Flask(__name__)

# Configure database - PythonAnywhere uses PostgreSQL
# For development, you can use SQLite
if 'PYTHONANYWHERE_DOMAIN' in os.environ:
    # Production database on PythonAnywhere
    app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///orders.db')
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
else:
    # Development database
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///orders.db'
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Model Database
class Order(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    order_number = db.Column(db.String(50), unique=True, nullable=False)
    date = db.Column(db.DateTime, nullable=False)
    rim_quantity = db.Column(db.Integer, nullable=False)
    city = db.Column(db.String(100), nullable=False)
    document_type = db.Column(db.String(100), nullable=False)
    unit_price = db.Column(db.Float, nullable=False)
    total_price = db.Column(db.Float, nullable=False)
    entry_date = db.Column(db.DateTime)
    print_deadline = db.Column(db.DateTime)
    cek_date = db.Column(db.DateTime)
    finish_date = db.Column(db.DateTime)
    status = db.Column(db.String(50), default='masuk')  # masuk, proses, cek, finish
    notes = db.Column(db.Text)

    def to_dict(self):
        return {
            'id': self.id,
            'order_number': self.order_number,
            'date': self.date.strftime('%Y-%m-%d') if self.date else None,
            'rim_quantity': self.rim_quantity,
            'city': self.city,
            'document_type': self.document_type,
            'unit_price': self.unit_price,
            'total_price': self.total_price,
            'entry_date': self.entry_date.strftime('%Y-%m-%d') if self.entry_date else None,
            'print_deadline': self.print_deadline.strftime('%Y-%m-%d') if self.print_deadline else None,
            'cek_date': self.cek_date.strftime('%Y-%m-%d') if self.cek_date else None,
            'finish_date': self.finish_date.strftime('%Y-%m-%d') if self.finish_date else None,
            'status': self.status,
            'notes': self.notes
        }

# Route
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/orders', methods=['GET'])
def get_orders():
    orders = Order.query.all()
    return jsonify([order.to_dict() for order in orders])

@app.route('/api/orders', methods=['POST'])
def add_order():
    data = request.json
    
    # Parse dates
    date = datetime.strptime(data['date'], '%Y-%m-%d') if data.get('date') else datetime.now()
    entry_date = datetime.strptime(data['entry_date'], '%Y-%m-%d') if data.get('entry_date') else None
    print_deadline = datetime.strptime(data['print_deadline'], '%Y-%m-%d') if data.get('print_deadline') else None
    
    order = Order(
        order_number=data['order_number'],
        date=date,
        rim_quantity=data['rim_quantity'],
        city=data['city'],
        document_type=data['document_type'],
        unit_price=data['unit_price'],
        total_price=data['unit_price'] * data['rim_quantity'],
        entry_date=entry_date,
        print_deadline=print_deadline,
        status=data.get('status', 'masuk'),
        notes=data.get('notes', '')
    )
    
    db.session.add(order)
    db.session.commit()
    return jsonify(order.to_dict()), 201

@app.route('/api/orders/<int:order_id>', methods=['PUT'])
def update_order(order_id):
    order = Order.query.get_or_404(order_id)
    data = request.json
    
    # Update basic fields
    order.order_number = data.get('order_number', order.order_number)
    order.rim_quantity = data.get('rim_quantity', order.rim_quantity)
    order.city = data.get('city', order.city)
    order.document_type = data.get('document_type', order.document_type)
    order.unit_price = data.get('unit_price', order.unit_price)
    order.total_price = order.rim_quantity * order.unit_price
    order.status = data.get('status', order.status)
    order.notes = data.get('notes', order.notes)
    
    # Parse and update dates
    if data.get('date'):
        order.date = datetime.strptime(data['date'], '%Y-%m-%d')
    if data.get('entry_date'):
        order.entry_date = datetime.strptime(data['entry_date'], '%Y-%m-%d')
    if data.get('print_deadline'):
        order.print_deadline = datetime.strptime(data['print_deadline'], '%Y-%m-%d')
    if data.get('cek_date'):
        order.cek_date = datetime.strptime(data['cek_date'], '%Y-%m-%d')
    if data.get('finish_date'):
        order.finish_date = datetime.strptime(data['finish_date'], '%Y-%m-%d')
    
    db.session.commit()
    return jsonify(order.to_dict())

@app.route('/api/orders/<int:order_id>', methods=['DELETE'])
def delete_order(order_id):
    order = Order.query.get_or_404(order_id)
    db.session.delete(order)
    db.session.commit()
    return '', 204

@app.route('/api/export/excel')
def export_to_excel():
    orders = Order.query.all()
    data = [order.to_dict() for order in orders]
    df = pd.DataFrame(data)
    
    # Create a bytes buffer for the Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Orders')
        
        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Orders']
        
        # Add some formatting
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Write the column headers with the defined format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Adjust column widths
        for i, col in enumerate(df.columns):
            column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_width)
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'orders_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )

# Create database tables
with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=True)