CREATE USER [ori19] FROM EXTERNAL PROVIDER;
ALTER ROLE db_datareader ADD MEMBER [ori19];
ALTER ROLE db_datawriter ADD MEMBER [ori19];
GRANT EXECUTE TO [ori19];

Select * from DeltaTokens