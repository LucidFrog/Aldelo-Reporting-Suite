class ReportController < ActionController::Base
  # Need these things to interact with .mdb files 
  # https://gist.github.com/407804 this post fixed a big bug with 1.9.2 and win32ole
  # Pretty huge
  
  
  require 'win32ole'
  require 'Win32API' 
  CoInitialize = Win32API.new('ole32', 'CoInitialize', 'P', 'L')

  # For Employee Report
  def all_employees
    CoInitialize.call( 0 )
    db = AccessDb.new(DBLOCATION)
    db.open
    
    # Query the DB here
    db.query("SELECT FirstName, LastName, JobTitleText, SocialSecurityNumber FROM EmployeeFiles, JobTitles WHERE EmployeeFiles.JobTitleID = JobTitles.JobTitleID;")
    @field_names = db.fields
    @all_employees = db.data
  end

  def export_all_employees
    @outfile = "employees_" + Time.now.strftime("%m-%d-%Y") + ".csv"
    send_data Report.all_employee_to_csv(), :type => 'text/csv; charset=iso-8859-1; header=present', :disposition => "attachment; filename=#{@outfile}"
  end

  # For Payroll Report
  def payroll
    #waiting on input, don't have to do anything yet.
    if params[:payroll_date]
      @payroll_data = Report.payroll_data(params[:payroll_date])
      @payroll_data.sort!
      @date = params[:payroll_date]
      
      render :partial => "payroll"
    end
  end

  def export_payroll
    if params[:payroll_date]
      date = params[:payroll_date]
      @outfile = "payroll_"+date+".csv"
      send_data Report.payroll_to_csv(date), :type => 'text/csv; charset=iso-8859-1; header=present', :disposition => "attachment; filename=#{@outfile}"
    end
  end
  
  # For Overtime Report
  def overtime
    CoInitialize.call( 0 )
    db = AccessDb.new(DBLOCATION)
    db.open
    # Query Here For Basic Selection Information
    
    db.query("SELECT JobTitleID, JobTitleText FROM JobTitles;")
    @job_titles = db.data

    #waiting on input, don't have to do anything yet.
    if params[:overtime_date] && params[:job_title_id]
      #Lets query! 
      CoInitialize.call( 0 )
      db = AccessDb.new(DBLOCATION)
      db.open

      #Get all of the SSNS for the Employees that have the selected JOBTITLE      
      db.query("SELECT SocialSecurityNumber FROM EmployeeFiles WHERE JobTitleID = "+params[:job_title_id]+";")
      @ssns = db.data

      #structure the last string to be attached to the query
      employee_ot_query = ""
      counter = 0
      for ssn in @ssns
        employee_ot_query += "EmployeeTimeCards.WorkDate >= #"+params[:overtime_date]+"# AND EmployeeFiles.SocialSecurityNumber = '"+ssn[0]+"'"
        counter +=1
        if counter > 0 && counter != @ssns.length
          employee_ot_query += " OR EmployeeTimeCards.TotalWeeklyOverTimeMinutes > 0 AND "
        end
      end 
      
      # Query Here
      db.query("SELECT EmployeeFiles.FirstName, EmployeeFiles.LastName, JobTitles.JobTitleText, EmployeeTimeCards.WorkDate, EmployeeTimeCards.TotalWeeklyOverTimeMinutes
        FROM (EmployeeFiles INNER JOIN JobTitles ON EmployeeFiles.JobTitleID = JobTitles.JobTitleID) 
        INNER JOIN EmployeeTimeCards ON EmployeeFiles.EmployeeID = EmployeeTimeCards.EmployeeID 
        WHERE EmployeeTimeCards.WorkDate > #"+params[:overtime_date]+"#
        AND EmployeeTimeCards.TotalWeeklyOverTimeMinutes > 0 AND "+employee_ot_query+";")
      
      @overtime_data = db.data
      @overtime_data.sort!
      @date = params[:overtime_date]
      @job_title_id = params[:job_title_id]
      
      render :partial => "overtime"
    end
  end
  
  def export_overtime
    if params[:overtime_date] && params[:job_title_id]
      date = params[:overtime_date]
      job_title_id = params[:job_title_id]
      @outfile = "overtime_"+date+".csv"
      send_data Report.overtime_to_csv(date, job_title_id), :type => 'text/csv; charset=iso-8859-1; header=present', :disposition => "attachment; filename=#{@outfile}"
    end
  end
  
  
  # For Overtime Report
  def total_hours
    CoInitialize.call( 0 )
    db = AccessDb.new(DBLOCATION)
    db.open
    # Query Here For Basic Selection Information
    
    db.query("SELECT JobTitleID, JobTitleText FROM JobTitles;")
    @job_titles = db.data

    #waiting on input, don't have to do anything yet.
    if params[:total_hours_date] && params[:job_title_id]
      #Lets query! 
      CoInitialize.call( 0 )
      db = AccessDb.new(DBLOCATION)
      db.open

      #Get all of the SSNS for the Employees that have the selected JOBTITLE      
      db.query("SELECT SocialSecurityNumber FROM EmployeeFiles WHERE JobTitleID = "+params[:job_title_id]+";")
      @ssns = db.data

      #structure the last string to be attached to the query
      employee_ot_query = ""
      counter = 0
      for ssn in @ssns
        employee_ot_query += "EmployeeTimeCards.WorkDate >= #"+params[:total_hours_date]+"# AND EmployeeFiles.SocialSecurityNumber = '"+ssn[0]+"'"
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
      
      @total_hours_data = db.data
      @total_hours_data.sort!
      @date = params[:total_hours_date]
      @job_title_id = params[:job_title_id]
      
      render :partial => "total_hours"
    end
  end
  
  def export_total_hours
    if params[:total_hours_date] && params[:job_title_id]
      date = params[:total_hours_date]
      job_title_id = params[:job_title_id]
      @outfile = "total_hours_"+date+".csv"
      send_data Report.total_hours_to_csv(date, job_title_id), :type => 'text/csv; charset=iso-8859-1; header=present', :disposition => "attachment; filename=#{@outfile}"
    end
  end
  
  
  
  
  
  # For Liquor Sales Report
  def liquor_sales
    #waiting on input, don't have to do anything yet.
    if params[:start_date] && params[:liquor_name] == '' || params[:start_date] && params[:liquor_name] == nil
      #Lets query 
      CoInitialize.call( 0 )
      db = AccessDb.new(DBLOCATION)
      db.open
      
      # Query Here        
      db.query("SELECT DISTINCT OrderTransactions.OrderTransactionID, OrderTransactions.OrderID, MenuItems.MenuItemText, OrderHeaders.OrderDateTime,
         EmployeeFiles.FirstName, EmployeeFiles.LastName,
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod1ID = MenuModifiers.MenuModifierID),
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod2ID = MenuModifiers.MenuModifierID),
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod3ID = MenuModifiers.MenuModifierID)
        FROM ((OrderTransactions INNER JOIN MenuItems ON OrderTransactions.MenuItemID = MenuItems.MenuItemID)
        INNER JOIN OrderHeaders ON OrderHeaders.OrderID = OrderTransactions.OrderID)
        INNER JOIN EmployeeFiles ON EmployeeFiles.EmployeeID = OrderHeaders.EmployeeID
        WHERE MenuItems.MenuItemText = '"+params[:liquor_type]+"'
        AND OrderHeaders.OrderDateTime > #"+params[:start_date].to_s+"#
        AND OrderHeaders.OrderDateTime <= #"+params[:end_date].to_s+"#;")

      @liquor_sales_data = db.data
      @liquor_sales_data.sort!
      @liquor_type = params[:liquor_type]
      @start_date = params[:start_date]
      @end_date = params[:end_date]

      render :partial => "liquor_sales"
    #Ok lets see if we have a liquor name
    elsif params[:start_date] && params[:liquor_name]
      #There is a liquor name, only search for results containing 
      #this liquor name in mod1 mod2 and mod3
      #Lets query 
      CoInitialize.call( 0 )
      db = AccessDb.new(DBLOCATION)
      db.open
      
      # Query Here        
      db.query("SELECT  OrderTransactions.OrderTransactionID, OrderTransactions.OrderID, MenuItems.MenuItemText, OrderHeaders.OrderDateTime,
         EmployeeFiles.FirstName, EmployeeFiles.LastName,
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod1ID = MenuModifiers.MenuModifierID),
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod2ID = MenuModifiers.MenuModifierID),
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod3ID = MenuModifiers.MenuModifierID)
        FROM ((OrderTransactions INNER JOIN MenuItems ON OrderTransactions.MenuItemID = MenuItems.MenuItemID)
        INNER JOIN OrderHeaders ON OrderHeaders.OrderID = OrderTransactions.OrderID)
        INNER JOIN EmployeeFiles ON EmployeeFiles.EmployeeID = OrderHeaders.EmployeeID
        WHERE MenuItems.MenuItemText = '"+params[:liquor_type]+"'
        AND OrderHeaders.OrderDateTime > #"+params[:start_date].to_s+"#
        AND OrderHeaders.OrderDateTime <= #"+params[:end_date].to_s+"#;")

      @liquor_sales = db.data
      @liquor_sales.sort!
      @liquor_type = params[:liquor_type]
      @liquor_name = params[:liquor_name]
      @start_date = params[:start_date]
      @end_date = params[:end_date]

      @liquor_sales_data = Array.new
      @liquor_sales.each do |sale|
        if sale[4] == params[:liquor_name] || sale[5] == params[:liquor_name] || sale[6] == params[:liquor_name]
          @liquor_sales_data << sale 
        end
      end

      @shots = 0
      @on_the_rocks = 0
      @mixed = 0

       p @liquor_sales_data

      for sale in @liquor_sales_data
        if sale[5] == 'Shot' || sale[6] == 'Shot' || sale[7] == 'Shot'
          @shots +=1
        elsif sale[5] == 'On The Rocks' || sale[6] == 'On The Rocks' || sale[7] == 'On The Rocks'
          @on_the_rocks +=1
        else 
          @mixed +=1
        end
      end

      render :partial => "liquor_sales_name"
    end
  end#end liquor_sales

  def export_liquor_sales
    if params[:start_date]
      start_date = params[:start_date]
      end_date = params[:end_date]
      liquor_type = params[:liquor_type]

      @outfile = liquor_type+"_"+start_date+"_to_"+end_date+".csv"
      send_data Report.liquor_sales_to_csv(start_date, end_date, liquor_type), :type => 'text/csv; charset=iso-8859-1; header=present', :disposition => "attachment; filename=#{@outfile}"
    end
  end
  
  def export_liquor_sales_name
    if params[:start_date]
      start_date = params[:start_date]
      end_date = params[:end_date]
      liquor_type = params[:liquor_type]
      liquor_name = params[:liquor_name]

      @outfile = liquor_name+"_"+start_date+"_to_"+end_date+".csv"
      send_data Report.liquor_sales_name_to_csv(start_date, end_date, liquor_type, liquor_name), :type => 'text/csv; charset=iso-8859-1; header=present', :disposition => "attachment; filename=#{@outfile}"
    end
  end
  
  
end#end Report Controller
