// Revealing Module pattern
// Immediately Invoked Function returns an object with methods attached.
var SPREADSHEET2SQL = (function() {
  var ss2sql = {}; // Object to be returned to which methods are attached
  
  /*
  Private function (not returned in object "ss2sql").
  Function that determines the type of the given value.
  The logic used here is crucial to how the application
  assigns data types to columns.
  Returns an array of types.
  */
  function getValueType(value) {
    // First check for date type
    if (!isNaN(Date.parse(value))) {
      return 'date';
    }
    // If numeric, distinguish between "float" and "integer".
    if (Number(parseFloat(value)) == value) {
      // Trick to convert value to integer (parseInt) and then
      // compare this to the actual value multiplied by 1000.
      // If the same, assign it as "integer", else call it a "float"
      // If you need more precision (e.g. input is 10.0002), then
      // increase the 1000 to 10000 to make this return "float"
      if (parseInt(value) * 1000 == value * 1000) {
        return 'integer';
      } else {
        return 'float';
      }
    }
    // If all previuos checks return "false", assume it is a string.
    return 'string';
  }
  /*
  Private function (not returned in object "ss2sql")
  Uses function "getValueType()" to determine all the types in an 
  array of values argument and return a list of the data types.
  */
  function getColumnTypes(values) {
    var valueType,
        columnTypes = [];
    for(var i = 0; i < values.length; i +=1) {
      // Skip empty values coming from empty cells
      if(typeof values[i] == 'undefined' || values[i] === '') {
        continue; 
      }      
      valueType = getValueType(values[i]);
      // Ensure that types are only added once to the array.
      if(columnTypes.indexOf(valueType) < 0) {
        columnTypes.push(valueType);
      }
    }
    return columnTypes;   
  }
  /*
  Private function (not returned in object "ss2sql").
  Takes an array of column types as returned by function "getColumnTypes()"
  and applies a series of tests to the array to assign a type for the column values.
  */
  function assignColumnType(columnTypes) {
    // If "string" is present then assign this type regardless of any other types present.
    if(columnTypes.indexOf('string') > -1) {
      return 'string';
    }
    // If the only type present is "date" then assign this type.
    if(columnTypes.length === 1 && columnTypes.indexOf('date') > -1) {
      return 'date'
    }
    // If the only type present is "float" then assign this type.
    if(columnTypes.length === 1 && columnTypes.indexOf('float') > -1) {
       return 'float';
    }
    //// If the only type present is "integer" then assign this type.
    if(columnTypes.length === 1 && columnTypes.indexOf('integer') > -1) {
       return 'integer';
    }
    // If the types "float" and "integer" are present, assign type "float".
    if(columnTypes.length === 2 && (columnTypes.indexOf('float') > -1 && columnTypes.indexOf('integer') > -1)) {
       return 'float';
    }
    // Default type
    return 'string';
  }
  // Public methods set below. These are attached to returned "ss2sql" object. This is
  // the "revealing" part of the pattern.
  /////////////////////////////////////////////////////////////////////////////////////
  /*
  Given a "Range" object as an argument, set a series of relevant instance variables
  using "Range" object methods.
  */
  ss2sql.setRange = function(rng) {
    this.rng = rng;
    this.rngAddr = this.rng.getA1Notation();
    this.rngSheet = this.rng.getSheet();
    this.rngSheetName = this.rngSheet.getName();
    this.rngRowCount = this.rng.getNumRows();
    this.rngColCount = this.rng.getNumColumns();
    this.rngValues = this.rng.getValues();
  };
  /*
  Set the column header row. If no argument is given, assume it is row 1.
  The column names used in the generated SQL output are taken from this row.
  This method sets a number of instance variables relating to the header row and the data rows.
  These are described in comments within the method.
  */
  ss2sql.setHeaderRow = function(headerRowNum) {
    var firstHeaderRowCellAddr,
        lastHeaderRowCellAddr,
        lastColNum = this.rng.getLastColumn();
    // If no row number is given for the header, assume it is row 1.
    if (typeof headerRowNum === 'undefined') {
      this.headerRowNum = 1;
    } else {
      this.headerRowNum = headerRowNum;
    }
    // Define the header row range.
    firstHeaderRowCellAddr = this.rng.getCell(this.headerRowNum, 1).getA1Notation();
    lastHeaderRowCellAddr = this.rng.getCell(this.headerRowNum, lastColNum).getA1Notation();
    this.headerRowRndAddr = firstHeaderRowCellAddr + ':' + lastHeaderRowCellAddr;
    this.headerRowRng = this.rngSheet.getRange(this.headerRowRndAddr);
    // Convert the header row values to lower case, trim leading and lagging white space
    // and convert non-alphanumeric characters to underscore.
    this.headerRowNames = this.headerRowRng.getValues()[0].map(function(element) {
      var normalizedElement = element.toLowerCase().trim().replace(/[^a-z0-9]/g, '_');
      return normalizedElement;
    });
    this.rngDataValues = this.rngValues.slice(this.headerRowNum, this.rngValues.length);
  },
  // A series of get methods for returning instance variables.
  ss2sql.getRngAddr = function() {
    return this.rngAddr;
  };
  ss2sql.getRngSheet = function() {
    return this.rngSheet;
  };
  ss2sql.getRngSheetName = function() {
    return this.rngSheetName;
  };
  ss2sql.getRngRowCount = function() {
    return this.rngRowCount;
  };
  ss2sql.getRngColCount = function() {
    return this.rngColCount;
  };
  ss2sql.getHeaderRowAddr = function() {
    return this.headerRowRndAddr;
  };
  ss2sql.getHeaderRowRng = function() {
    return this.headerRowRng;
  };
  ss2sql.getHeaderRowNames = function() {
    return this.headerRowNames;
  };
  ss2sql.getRngDataValues = function() {
    return this.rngDataValues;
  };
  /* Return an object mapping column names (derived from header row) to 
   an array the data values extracted from the data rows.
   */
  ss2sql.getColumnNameValuesMap = function() {
    var columnNameValuesMap = {},
        values = [],
        colNum,
        firstRowAddr,
        lastRowAddr,
        i,
        colName;
    for(colNum = 0; colNum < this.headerRowNames.length; colNum +=1) {
      colName = this.headerRowNames[colNum];
      columnNameValuesMap[colName] = [];
      for(var i = 0; i < this.rngDataValues.length; i +=1) {
        columnNameValuesMap[colName].push(this.rngDataValues[i][colNum]);
      }
    }
    return columnNameValuesMap;
  };
  /* Return an object mapping the assigned column names to the assigned data types.
  */
  ss2sql.getColumnNameTypeMap = function() {
    var columnNameValuesMap = this.getColumnNameValuesMap(),
        columnNameTypeMap = {},
        colNum;
    for(colNum = 0; colNum < this.headerRowNames.length; colNum +=1) {
      columnNameTypeMap[this.headerRowNames[colNum]] = assignColumnType(getColumnTypes(columnNameValuesMap[this.headerRowNames[colNum]]));
    }
    return columnNameTypeMap;
  };
  /*
  Return the "CREATE TABLE" SQL statement for the table name and RDBMS arguments.
  */
  ss2sql.makeCreateTableSql = function(tableName, dbms) {
    var colNum,
        columnName,
        columnType,
        pgColumnType,
        columnDefn,
        columnNameTypeMap = this.getColumnNameTypeMap(),
        ddl,
        pkColumnName,
        // This nested object maps the data types defined in the code here ("string", "date", "float" and "integer" to
        // their equivalents for each of the three supported RDBMSs (PostgreSQL, MySQL and SQLite).
        dbms_type_map = {postgres: {primary_key: 'SERIAL PRIMARY KEY', string: 'CHARACTER VARYING', date: 'DATE', float: 'NUMERIC', integer: 'INTEGER'},
                       sqlite:{primary_key: 'INTEGER PRIMARY KEY', string: 'TEXT', date: 'DATE', float: 'REAL', integer: 'INTEGER'},
                         mysql: {primary_key: 'INTEGER NOT NULL PRIMARY KEY AUTO_INCREMENT', string: 'VARCHAR(4000)', date: 'DATE', float: 'DECIMAL(11,5)'}};
    pkColumnName = tableName.replace(/\./, '_') + '_id';
    // Loop over all the column names and build up the "CREATE TABLE" statement.
    ddl = 'CREATE TABLE ' + tableName + '(\n  ' + pkColumnName + ' ' + dbms_type_map[dbms]['primary_key'] + ',\n';
    for(colNum = 0; colNum < this.headerRowNames.length; colNum +=1) {
      columnName = this.headerRowNames[colNum];
      columnType = columnNameTypeMap[columnName];
      if(columnType === 'string') {
        pgColumnType = dbms_type_map[dbms]['string'];
      } else if(columnType === 'date') {
        pgColumnType = dbms_type_map[dbms]['date'];
      } else if(columnType === 'float') {
        pgColumnType = dbms_type_map[dbms]['float'];
      } else if(columnType === 'integer') {
        pgColumnType = dbms_type_map[dbms]['integer'];
      }
      columnDefn = '  ' + columnName + ' ' + pgColumnType + ',\n';
      ddl += columnDefn;
    }
    // Delete the trailing comma left by the loop above and add the terminating ");" string
    ddl = ddl.replace(/,\s*$/, ');\n');
    return ddl;
  };
  /*
  Create the "INSERT" SQL statements, one for each data row.
  Return them as an array of statements.
  */
  ss2sql.makeInsertSql = function(tableName) {
    var sql,
        columnNameTypeMap,
        insertStatements = [],
        colNum,
        columnName,
        columnType,
        dataRowNum,
        rowValuesIn = [],
        rowValuesOut = [],
        rowValuesInsert = [],
        rowValue;
    sql  = 'INSERT INTO ' + tableName + '(' + this.headerRowNames.join(',') + ') VALUES(';
    columnNameTypeMap = this.getColumnNameTypeMap();
    for(dataRowNum = 0; dataRowNum < this.rngDataValues.length; dataRowNum +=1) {
      rowValuesIn = this.rngDataValues[dataRowNum];
      rowValuesOut = [];
      for(colNum = 0; colNum < this.headerRowNames.length; colNum +=1) {
        columnName = this.headerRowNames[colNum];
        columnType = columnNameTypeMap[columnName];
        rowValue = this.rngDataValues[dataRowNum][colNum];
        // Empty values are assigned NULL in the "INSERT" statements
        if(rowValue.toString().length < 1) {
          rowValue = 'NULL';
          rowValuesOut.push(rowValue);
          continue;
        }
        if(columnType === 'string') {
          // Deals with embedded single quotes in values by replacing them with two single quotes
          rowValue = rowValue.toString().replace(/[']/g, "''");
          rowValue = "'" + rowValue + "'";
        }
        if(columnType === 'date') {
          // Parentheses are required around "rowValue.getMonth() + 1" to prevent string concatenation.
          rowValue = rowValue.getFullYear() + '-' + (rowValue.getMonth() + 1) + '-' + rowValue.getDate();
          rowValue = "'" + rowValue + "'";
        }
        rowValuesOut.push(rowValue);
      }
      rowValuesInsert.push(sql + rowValuesOut.join(',') + ');');
    }
    return rowValuesInsert;
  };
  return ss2sql;
}());
