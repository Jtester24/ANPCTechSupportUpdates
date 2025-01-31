from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
import logging

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data.db'  # Example with SQLite
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# Set up logging
logging.basicConfig(level=logging.DEBUG)

class ClientData(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    site = db.Column(db.String(100))
    timestamp = db.Column(db.String(100))
    name = db.Column(db.String(100))
    email = db.Column(db.String(100))
    phone = db.Column(db.String(100))
    description = db.Column(db.Text)
    guidance_interrupted = db.Column(db.Boolean)
    system_guidance_available = db.Column(db.Boolean)
    part_replacement_required = db.Column(db.Boolean)
    network_info = db.Column(db.Text)  # Store network info as JSON or Text

@app.route('/submit', methods=['POST'])
def submit():
    try:
        data = request.get_json()
        new_client = ClientData(
            site=data['site'],
            timestamp=data['timestamp'],
            name=data['name'],
            email=data['email'],
            phone=data['phone'],
            description=data['description'],
            guidance_interrupted=data['guidance_interrupted'],
            system_guidance_available=data['system_guidance_available'],
            part_replacement_required=data['part_replacement_required'],
            network_info=data['network_info']
        )
        db.session.add(new_client)
        db.session.commit()
        logging.info("New client added successfully")
        return jsonify({'message': 'New client added successfully'})
    except Exception as e:
        logging.error(f"Error occurred: {e}")
        return jsonify({'error': 'An error occurred'}), 500

if __name__ == '__main__':
    try:
        logging.info("Creating database tables...")
        with app.app_context():
            db.create_all()  # Create database tables
        logging.info("Starting Flask server...")
        app.run(debug=True)
    except Exception as e:
        logging.error(f"Failed to start server: {e}")