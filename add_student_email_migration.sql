CREATE DATABASE IF NOT EXISTS libratrack_db
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE libratrack_db;

CREATE TABLE IF NOT EXISTS users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(50) NOT NULL UNIQUE,
    full_name VARCHAR(100) NOT NULL,
    password_hash CHAR(64) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS students (
    student_id VARCHAR(50) PRIMARY KEY,
    barcode_value VARCHAR(50) NOT NULL UNIQUE,
    full_name VARCHAR(150) NOT NULL,
    email VARCHAR(150) NOT NULL,
    age INT NOT NULL,
    year_level VARCHAR(50) DEFAULT '',
    course VARCHAR(100) DEFAULT '',
    address TEXT,
    last_attendance DATETIME NULL,
    status VARCHAR(30) NOT NULL DEFAULT 'Registered',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS attendance (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id VARCHAR(50) NOT NULL,
    attendance_date DATE NOT NULL,
    time_in DATETIME NULL,
    time_out DATETIME NULL,
    last_action VARCHAR(20) NOT NULL DEFAULT 'TIME IN',
    status VARCHAR(20) NOT NULL DEFAULT 'OPEN',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_attendance_student
        FOREIGN KEY (student_id) REFERENCES students(student_id)
        ON DELETE CASCADE
);

CREATE INDEX idx_attendance_student_date
    ON attendance (student_id, attendance_date);

CREATE INDEX idx_attendance_date
    ON attendance (attendance_date);

CREATE TABLE IF NOT EXISTS books (
    id INT AUTO_INCREMENT PRIMARY KEY,
    book_id VARCHAR(50) NOT NULL UNIQUE,
    barcode_value VARCHAR(80) NOT NULL UNIQUE,
    title VARCHAR(255) NOT NULL,
    author VARCHAR(150) DEFAULT '',
    category VARCHAR(100) DEFAULT '',
    status VARCHAR(20) NOT NULL DEFAULT 'AVAILABLE',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS borrow_transactions (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id VARCHAR(50) NOT NULL,
    book_id INT NOT NULL,
    borrow_date DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    due_date DATETIME NOT NULL,
    return_date DATETIME NULL,
    status VARCHAR(20) NOT NULL DEFAULT 'BORROWED',
    reminder_sent TINYINT(1) NOT NULL DEFAULT 0,
    overdue_notice_sent TINYINT(1) NOT NULL DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    CONSTRAINT fk_borrow_student
        FOREIGN KEY (student_id) REFERENCES students(student_id)
        ON DELETE CASCADE,
    CONSTRAINT fk_borrow_book
        FOREIGN KEY (book_id) REFERENCES books(id)
        ON DELETE CASCADE
);

CREATE INDEX idx_books_book_id ON books (book_id);
CREATE INDEX idx_books_barcode ON books (barcode_value);
CREATE INDEX idx_borrow_active ON borrow_transactions (status, due_date);
CREATE INDEX idx_borrow_student_status ON borrow_transactions (student_id, status);
CREATE INDEX idx_borrow_book_status ON borrow_transactions (book_id, status);

INSERT INTO users (username, full_name, password_hash)
VALUES ('admin', 'System Administrator', SHA2('admin123', 256))
ON DUPLICATE KEY UPDATE
    full_name = VALUES(full_name),
    password_hash = VALUES(password_hash);

ALTER TABLE borrow_transactions
    ADD COLUMN IF NOT EXISTS overdue_notice_sent TINYINT(1) NOT NULL DEFAULT 0;
