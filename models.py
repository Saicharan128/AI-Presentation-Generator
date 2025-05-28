from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class FileRecord(db.Model):
    __tablename__ = 'file_records'
    
    id = db.Column(db.Integer, primary_key=True)
    user_email = db.Column(db.String(255), nullable=False, index=True)
    topic = db.Column(db.String(500), nullable=False)
    file_type = db.Column(db.String(10), nullable=False)
    file_data = db.Column(db.Text, nullable=False)
    yaml_content = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    
    # Add unique constraint to prevent duplicates
    __table_args__ = (
        db.UniqueConstraint('user_email', 'topic', 'file_type', name='uq_user_topic_type'),
    )
    
    def __repr__(self):
        return f'<FileRecord {self.id}: {self.topic} - {self.file_type}>'
    
    def __init__(self, user_email, topic, file_type, file_data, yaml_content=None):
        self.user_email = user_email
        self.topic = topic
        self.file_type = file_type
        self.file_data = file_data
        self.yaml_content = yaml_content