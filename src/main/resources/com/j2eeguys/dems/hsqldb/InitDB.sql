CREATE TABLE WORKER (
    id int identity primary key,
    VR_ID varchar(25) NOT NULL,
    LAST_NAME varchar(64) NOT NULL,
    FIRST_NAME varchar(64) NOT NULL,
    PRECINT SMALLINT DEFAULT NULL,
    ROLE varchar(255) DEFAULT NULL,
);

CREATE TABLE AVAILABILITY (
    id int NOT NULL,
    DAY DATE NOT NULL,
);
