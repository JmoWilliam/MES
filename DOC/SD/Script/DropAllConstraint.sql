  BEGIN
 
  DECLARE @stmt VARCHAR(300);
 
  -- Cursor to generate ALTER TABLE DROP CONSTRAINT statements  
  DECLARE cur CURSOR FOR
     SELECT 'ALTER TABLE ' + OBJECT_SCHEMA_NAME(parent_object_id) + '.[' + OBJECT_NAME(parent_object_id) +
                    '] DROP CONSTRAINT ' + name
     FROM sys.foreign_keys ;
 
   OPEN cur;
   FETCH cur INTO @stmt;
 
   -- Drop each found foreign key constraint 
   WHILE @@FETCH_STATUS = 0
     BEGIN
       EXEC (@stmt);
       FETCH cur INTO @stmt;
     END
 
  CLOSE cur;
  DEALLOCATE cur;
 
  END
  GO