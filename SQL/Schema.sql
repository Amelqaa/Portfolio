#Create schema called SnubHub
CREATE SCHEMA SnubHub;

#Switch to a SnubHub schema
USE SnubHub;

#Create a table called Users in SnubHub schema
CREATE TABLE Users
(
	UserID INT AUTO_INCREMENT PRIMARY KEY,
    UserLastName VARCHAR (20),
    UserFirstName VARCHAR (20),
    UserEmailAddress VARCHAR (40)
)
;

#Create a table called Venues in SnubHub schema
CREATE TABLE Venues
(
	VenueID INT AUTO_INCREMENT PRIMARY KEY,
	VenueName VARCHAR(25),
	VenueCity VARCHAR(25),
	VenueState CHAR(2)
)
;

#Create a table called EventsTable in SnubHub schema
CREATE TABLE EventsTable
(
	EventID INT AUTO_INCREMENT PRIMARY KEY,
    EventName VARCHAR (50),
    EventDateTime DATETIME,
    VenueID_FK INT,
    FOREIGN KEY (VenueID_FK) REFERENCES Venues (VenueID)
)
;

#Create a table called TicketListings in SnubHub schema
CREATE TABLE TicketListings
(
	TicketListingID INT AUTO_INCREMENT PRIMARY KEY,
	NumTickets INT,
	PricePerTicket DECIMAL (6,2),
    EventID_FK INT,
    SellerID_FK INT,
	FOREIGN KEY (EventID_FK) REFERENCES EventsTable (EventID),
    FOREIGN KEY (SellerID_FK) REFERENCES Users (UserID)
)
;

#Create a table called TicketSales in SnubHub schema
CREATE TABLE TicketSales
(
	TransactionID INT AUTO_INCREMENT PRIMARY KEY,
	BuyerID_FK INT,
	TicketListingID_FK INT,
	FOREIGN KEY (BuyerID_FK) REFERENCES Users(UserID),
	FOREIGN KEY (TicketListingID_FK) REFERENCES TicketListings (TicketListingID)
)
;

#Add a record to table Users
INSERT INTO Users (UserLastName, UserFirstName, UserEmailAddress)
VALUES
	('Andersen', 'Ella', 'ella.andersen@123.com'),
	('Bruno', 'Matt', 'Bruno.matt@123.com'),
	('Craft', 'Madeline', 'craft.madeline@123.com')
;


#Add a record to table Venues
INSERT INTO Venues (VenueName, VenueCity, VenueState)
VALUES
	('The Fillmore', 'Philadelphia', 'PA'),
    ('The Met', 'Philadelphia', 'PA'),
    ('The Union League', 'Philadelphia', 'PA'),
    ('The Dr. Mendoza Union', 'Philadelphia', 'PA'),
    ('The Pooh Shiesty Stadium', 'Philadelphia', 'PA')
;



#Add a record to table EventsTable
INSERT INTO EventsTable (EventName, EventDateTime, VenueID_FK)
VALUES
	('The Dr. Mendoza Induction', '2026-04-10 19:30:00', 3),
    ('Philly Jazz Night', '2026-04-18 18:00:00', 1),
    ('The Dr. Mendoza Soccer Classic', '2026-06-06 09:00:00', 5),
    ('Black Sabbath Conference', '2026-05-15 17:30:00', 4),
    ('Greta Van Fleet Summit', '2026-05-02 20:00:00', 2)
;



#Add a record to table TicketListings
INSERT INTO TicketListings (NumTickets, PricePerTicket, EventID_FK, SellerID_FK)
VALUES
	(500, 100.00 , 1, 1),
    (50, 40.00 , 2, 2),
    (2000, 200.00 , 3, 3),
    (100, 70.00 , 4, 3),
    (300, 120.00 , 5, 2)
;


#Add a record to table TicketSales
INSERT INTO TicketSales (BuyerID_FK, TicketListingID_FK)
VALUES
	(2, 1),
	(3, 2),
	(1, 3),
	(2, 4),
	(3, 5)
;


#Check data was inserted properly
SELECT * FROM Users;

#Check data was inserted properly
SELECT * FROM Venues;

#Check data was inserted properly
SELECT * FROM EventsTable;

#Check data was inserted properly
SELECT * FROM TicketListings;

#Check data was inserted properly
SELECT * FROM TicketSales;

#End of script
