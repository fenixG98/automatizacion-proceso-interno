CREATE DATABASE clients_E;

CREATE TABLE clients_E.clients_E (
	id INT auto_increment NOT NULL,
	client varchar(100) NOT NULL,
	E varchar(10) NOT NULL,
	primary key (id)
)
ENGINE=InnoDB
DEFAULT CHARSET=utf8mb4
COLLATE=utf8mb4_0900_ai_ci;
