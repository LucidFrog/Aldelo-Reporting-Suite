class Report < ActiveRecord::Base
  require 'csv'

  def self.all_employee_to_csv()
    db = SQLite3::Database.new("db/flanders.sqlite3")
    columns, *rows = db.execute("SELECT FirstName, LastName, JobTitleText, SocialSecurityNumber FROM EmployeeFiles, JobTitles WHERE EmployeeFiles.JobTitleID = JobTitles.JobTitleID;")
    @field_names = columns
    @all_employees = rows
    @all_employees.sort!
    
    data = CSV.generate(:force_quotes => true) do |row|
      row << ['FirstName', 'LastName', 'SSN', 'JobTitle']
      @all_employees.each do |employee|
        row << [employee[0].to_s,
            employee[1].to_s,
            employee[3].to_s,
            employee[2].to_s]
      end
    end
    return data
  end

  def self.payroll_data(date)
    #Lets query 
    db = SQLite3::Database.new("db/flanders.sqlite3")
    rows = db.execute("SELECT EmployeeFiles.FirstName, EmployeeFiles.LastName, JobTitles.JobTitleText, 
                       EmployeeFiles.SocialSecurityNumber, EmployeePayrollHistory.PayRate,
                       EmployeePayrollHistory.RegularHours, EmployeePayrollHistory.OTPayRate, 
                       EmployeePayrollHistory.OverTimeHours, EmployeePayrollHistory.AdditionalPay, 
                       EmployeePayrollHistory.TotalTips, EmployeePayrollHistory.PayPeriodEndDate 
                       FROM ((EmployeePayrollHistory INNER JOIN EmployeeFiles ON  
                       EmployeePayrollHistory.EmployeeID = EmployeeFiles.EmployeeID ) INNER JOIN JobTitles 
                       ON  EmployeeFiles.JobTitleID = JobTitles.JobTitleID) 
                       WHERE PayPeriodEndDate = ##{date}#; ")
    return rows
  end

  def self.payroll_to_csv(date)
    @payroll = Report.payroll_data(date)
    @payroll.sort!
    
    data = CSV.generate(:force_quotes => true) do |row|
      row << ['FirstName', 'LastName',  'JobTitleText',  'SocialSecurityNumber',  'PayRate', 'RegularHours',  'OTPayRate', 'OverTimeHours', 'PayableTips', 'NonPayableTips']
      @payroll.each do |employee|
        row << [employee[0],
          employee[1],
          employee[2],
          employee[3],
          sprintf( "%.02f" , employee[4] ),
          sprintf( "%.02f" , employee[5] ), 
          sprintf( "%.02f" , employee[6] ),
          sprintf( "%.02f" , employee[7] ),
          sprintf( "%.02f" , employee[8] ), 
          sprintf( "%.02f" , employee[9] )]
      end
    end
    return data
  end   

  def self.overtime_to_csv(date, jobTitleId)
    #Lets query! 
    CoInitialize.call( 0 )
    db = AccessDb.new(DBLOCATION)
    db.open

    #Get all of the SSNS for the Employees that have the selected JOBTITLE      
    db.query("SELECT SocialSecurityNumber FROM EmployeeFiles WHERE JobTitleID = "+jobTitleId+";")
    @ssns = db.data

    #structure the last string to be attached to the query
    employee_ot_query = ""
    counter = 0
    for ssn in @ssns
      employee_ot_query += "EmployeeTimeCards.WorkDate >= #"+date+"# AND EmployeeFiles.SocialSecurityNumber = '"+ssn[0]+"'"
      counter +=1
      if counter > 0 && counter != @ssns.length
        employee_ot_query += " OR EmployeeTimeCards.TotalWeeklyOverTimeMinutes > 0 AND "
      end
    end 
    
    # Query Here
    db.query("SELECT EmployeeFiles.FirstName, EmployeeFiles.LastName, JobTitles.JobTitleText, EmployeeTimeCards.WorkDate, EmployeeTimeCards.TotalWeeklyOverTimeMinutes
      FROM (EmployeeFiles INNER JOIN JobTitles ON EmployeeFiles.JobTitleID = JobTitles.JobTitleID) 
      INNER JOIN EmployeeTimeCards ON EmployeeFiles.EmployeeID = EmployeeTimeCards.EmployeeID 
      WHERE EmployeeTimeCards.WorkDate > #"+date+"#
      AND EmployeeTimeCards.TotalWeeklyOverTimeMinutes > 0 AND "+employee_ot_query+";")
    
    @overtime_data = db.data
    @overtime_data.sort!
    
    data = CSV.generate(:force_quotes => true) do |row|
      row << ['FirstName', 'LastName',  'JobTitleText',  'DateOfOvertime',  'OTHours']
      @overtime_data.each do |employee|
        row << [employee[0],
          employee[1],
          employee[2],
          Time.at(employee[3].to_i).strftime("%m/%d/%Y"),
          sprintf( "%.02f" , employee[4].to_f/60 )]
      end
    end
    return data
  end 
  
  
  def self.total_hours_to_csv(date, jobTitleId)
    #Lets query! 
    CoInitialize.call( 0 )
    db = AccessDb.new(DBLOCATION)
    db.open

    #Get all of the SSNS for the Employees that have the selected JOBTITLE      
    db.query("SELECT SocialSecurityNumber FROM EmployeeFiles WHERE JobTitleID = "+jobTitleId+";")
    @ssns = db.data

    #structure the last string to be attached to the query
    employee_ot_query = ""
    counter = 0
    for ssn in @ssns
      employee_ot_query += "EmployeeTimeCards.WorkDate >= #"+date+"# AND EmployeeFiles.SocialSecurityNumber = '"+ssn[0]+"'"
      counter +=1
      if counter > 0 && counter != @ssns.length
        employee_ot_query += " OR "
      end
    end 
    
    # Query Here
    db.query("SELECT EmployeeFiles.FirstName, EmployeeFiles.LastName, JobTitles.JobTitleText, EmployeeTimeCards.WorkDate, 
      EmployeeTimeCards.TotalWeeklyOverTimeMinutes, EmployeeTimeCards.TotalRegularMinutes, EmployeeTimeCards.TotalWorkMinutes
      FROM (EmployeeFiles INNER JOIN JobTitles ON EmployeeFiles.JobTitleID = JobTitles.JobTitleID) 
      INNER JOIN EmployeeTimeCards ON EmployeeFiles.EmployeeID = EmployeeTimeCards.EmployeeID 
      WHERE "+employee_ot_query+";")
      
    
    @overtime_data = db.data
    @overtime_data.sort!
    
    data = CSV.generate(:force_quotes => true) do |row|
      row << ['FirstName', 'LastName',  'JobTitleText',  'WorkDate',  'OTHours', 'RegularHours', 'TotalHours']
      @overtime_data.each do |employee|
        row << [employee[0],
          employee[1],
          employee[2],
          Time.at(employee[3].to_i).strftime("%m/%d/%Y"),
          sprintf( "%.02f" , employee[4].to_f/60 ),
          sprintf( "%.02f" , employee[5].to_f/60 ),
          sprintf( "%.02f" , employee[6].to_f/60 )]
      end
    end
    return data
  end 





  def self.liquor_sales_to_csv(start_date, end_date, liquor_type)
    CoInitialize.call( 0 )
    db = AccessDb.new(DBLOCATION)
    db.open
    
    #Query HERE
    db.query("SELECT  OrderTransactions.OrderTransactionID, OrderTransactions.OrderID, MenuItems.MenuItemText, OrderHeaders.OrderDateTime,
       EmployeeFiles.FirstName, EmployeeFiles.LastName,
      (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod1ID = MenuModifiers.MenuModifierID),
      (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod2ID = MenuModifiers.MenuModifierID),
      (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod3ID = MenuModifiers.MenuModifierID)
      FROM ((OrderTransactions INNER JOIN MenuItems ON OrderTransactions.MenuItemID = MenuItems.MenuItemID)
      INNER JOIN OrderHeaders ON OrderHeaders.OrderID = OrderTransactions.OrderID)
      INNER JOIN EmployeeFiles ON EmployeeFiles.EmployeeID = OrderHeaders.EmployeeID
      WHERE MenuItems.MenuItemText = '"+liquor_type+"'
      AND OrderHeaders.OrderDateTime > #"+start_date+"#
      AND OrderHeaders.OrderDateTime <= #"+end_date+"#;")
       
    @liquor_sales = db.data
    @liquor_sales.sort!
    
    data = CSV.generate(:force_quotes => true) do |row|
      row << ['FieldOne', 'FieldTwo',  'FieldThree',  'Date', 'FirstName', 'LastName', 'Order#']
      @liquor_sales.each do |sale|
        row << [sale[6],
          sale[7],
          sale[8],
          Time.at(sale[3].to_i).strftime("%m/%d/%Y - %I:%M%p"),
          sale[4],
          sale[5],
          sale[1]]
      end
    end
    return data
  end 

  def self.liquor_sales_name_to_csv(start_date, end_date, liquor_type, liquor_name)
    CoInitialize.call( 0 )
    db = AccessDb.new(DBLOCATION)
    db.open
    
    #Query HERE
    db.query("SELECT  OrderTransactions.OrderTransactionID, OrderTransactions.OrderID, MenuItems.MenuItemText, OrderHeaders.OrderDateTime,
       EmployeeFiles.FirstName, EmployeeFiles.LastName, 
      (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod1ID = MenuModifiers.MenuModifierID),
      (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod2ID = MenuModifiers.MenuModifierID),
      (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod3ID = MenuModifiers.MenuModifierID)
      FROM ((OrderTransactions INNER JOIN MenuItems ON OrderTransactions.MenuItemID = MenuItems.MenuItemID)
      INNER JOIN OrderHeaders ON OrderHeaders.OrderID = OrderTransactions.OrderID)
      INNER JOIN EmployeeFiles ON EmployeeFiles.EmployeeID = OrderHeaders.EmployeeID
      WHERE MenuItems.MenuItemText = '"+liquor_type+"'
      AND OrderHeaders.OrderDateTime > #"+start_date+"#
      AND OrderHeaders.OrderDateTime <= #"+end_date+"#;")
       
    @liquor_sales = db.data
    @liquor_sales.sort!
    
    @liquor_sales_data = Array.new
    @liquor_sales.each do |sale|
      if sale[6] == liquor_name || sale[7] == :liquor_name || sale[8] == :liquor_name
        @liquor_sales_data << sale 
      end
    end
    
    data = CSV.generate(:force_quotes => true) do |row|
      row << ['FieldOne', 'FieldTwo',  'FieldThree',  'Date', 'FirstName', 'LastName', 'Order#']
      @liquor_sales_data.each do |sale|
        row << [sale[6],
          sale[7],
          sale[8],
          Time.at(sale[3].to_i).strftime("%m/%d/%Y - %I:%M%p"),
          sale[4],
          sale[5],
          sale[1]]
      end
    end
    return data
  end 


end#end Class Report
