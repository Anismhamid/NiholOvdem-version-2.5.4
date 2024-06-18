import logging
from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.exc import SQLAlchemyError
from datetime import datetime

app = Flask(__name__)

app.logger.setLevel(logging.ERROR)

app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root:root@localhost:10011/local'
db = SQLAlchemy(app)

class Workers(db.Model):
    Rowid = db.Column(db.Integer, primary_key=True)
    hours = db.Column(db.String(20))
    taken = db.Column(db.String(100))
    wage = db.Column(db.String(80))
    address = db.Column(db.String(100))
    managername = db.Column(db.String(100))
    compname = db.Column(db.String(100))
    date = db.Column(db.String(50))
    worker_id = db.Column(db.String(11))
    phone = db.Column(db.String(20))
    workername = db.Column(db.String(100))

@app.route('/get', methods=['GET'])
def get_api():
    try:
        # Query all records from the Workers table
        workers = Workers.query.order_by(Workers.workername).all()
        # Serialize the data to JSON
        data = [{
            'Rowid': worker.Rowid,
            'hours': worker.hours,
            'taken': worker.taken,
            'wage': worker.wage,
            'address': worker.address,
            'managername': worker.managername,
            'compname': worker.compname,
            'date': worker.date,
            'worker_id': worker.worker_id,
            'phone': worker.phone,
            'workername': worker.workername,
        } for worker in workers]

        # Return the data as JSON
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500


@app.route('/post', methods=['POST'])
def post_api():
    data = request.get_json()
    current_date = datetime.now().strftime('%y-%m-%d')
    workers = Workers(
        workername=data.get('WORKERS_NAME_var', ''),
        phone=data.get('WORKERS_PHONE_var', ''),
        worker_id=data.get('WORKERS_ID_var', ''),
        compname=data.get('WORKERS_COMPANY_NAME_var', ''),
        managername=data.get('WORKERS_COMPANY_MANAGER_NAME_var', ''),
        address=data.get('WORKERS_WORK_ADDRESS_var', ''),
        date=current_date,
        wage=data.get('WORKERS_WAGE_var', ''),
        taken=data.get('WORKERS_Taken_var', ''),
        hours=data.get('WORKERS_Hours_Var', ''),
    )
    db.session.add(workers)
    db.session.commit()
    return jsonify({'message': 'Data added successfully'}), 201


@app.route('/update', methods=['PUT'])
def update_api():
    try:
        # Retrieve data from the request body
        data = request.get_json()

        # Update data using parameterized query
        Row_id = data.get('Rowid')
        worker = Workers.query.filter_by(Rowid=Row_id).first()  # Corrected column name
        if worker:
            worker.hours = data.get('hours')
            worker.taken = data.get('taken')
            worker.wage = data.get('wage')
            worker.address = data.get('address')
            worker.managername = data.get('managername')
            worker.compname = data.get('compname')
            # worker.date = data.get('date')
            worker.worker_id = data.get('worker_id')  # Corrected assignment of worker_id
            worker.phone = data.get('phone')
            worker.workername = data.get('workername')

            # Save the changes to the database
            db.session.commit()

            return jsonify({'message': 'Data updated successfully'}), 200
        else:
            return jsonify({'error': 'Worker not found'}), 404
    except SQLAlchemyError as e:
        db.session.rollback()
        return jsonify({'error': f'Database error: {str(e)}'}), 500
    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.route('/delete/<int:Rowid>', methods=['DELETE'])
def delete_api(Rowid):
    try:
        # البحث عن العامل في قاعدة البيانات
        worker = Workers.query.get(Rowid)
        if worker:
            # حذف العامل
            db.session.delete(worker)
            db.session.commit()

            return jsonify({'message': 'Data deleted successfully'}), 200
        else:
            return jsonify({'error': 'Worker not found'}), 404
    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True)
