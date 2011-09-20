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
    db = AccessDb.new('c:\flanders.mdb')
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
      #Lets query 
      CoInitialize.call( 0 )
      db = AccessDb.new('c:\flanders.mdb')
      db.open
      
      # Query Here
      db.query("SELECT EmployeeFiles.FirstName, EmployeeFiles.LastName, JobTitles.JobTitleText, 
        EmployeeFiles.SocialSecurityNumber, EmployeePayrollHistory.PayRate, EmployeePayrollHistory.RegularHours, 
        EmployeePayrollHistory.OTPayRate, EmployeePayrollHistory.OverTimeHours, EmployeePayrollHistory.AdditionalPay, 
        EmployeePayrollHistory.TotalTips, EmployeePayrollHistory.PayPeriodEndDate 
        FROM EmployeePayrollHistory, EmployeeFiles, JobTitles 
        WHERE EmployeePayrollHistory.EmployeeID = EmployeeFiles.JobTitleID 
        AND EmployeeFiles.JobTitleID = JobTitles.JobTitleID
        AND PayPeriodEndDate = #"+params[:payroll_date]+"#;")

      @payroll_data = db.data
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
    db = AccessDb.new('c:\flanders.mdb')
    db.open
    # Query Here For Basic Selection Information
    
    db.query("SELECT JobTitleID, JobTitleText FROM JobTitles;")
    @job_titles = db.data

    #waiting on input, don't have to do anything yet.
    if params[:overtime_date] && params[:job_title_id]
      #Lets query! 
      CoInitialize.call( 0 )
      db = AccessDb.new('c:\flanders.mdb')
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
    db = AccessDb.new('c:\flanders.mdb')
    db.open
    # Query Here For Basic Selection Information
    
    db.query("SELECT JobTitleID, JobTitleText FROM JobTitles;")
    @job_titles = db.data

    #waiting on input, don't have to do anything yet.
    if params[:total_hours_date] && params[:job_title_id]
      #Lets query! 
      CoInitialize.call( 0 )
      db = AccessDb.new('c:\flanders.mdb')
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
    if params[:start_date]
      #Lets query 
      CoInitialize.call( 0 )
      db = AccessDb.new('c:\deagle.mdb')
      db.open
      
      # Query Here
      db.query("SELECT  OrderTransactions.OrderTransactionID, MenuItems.MenuItemText, OrderHeaders.OrderDateTime,
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod1ID = MenuModifiers.MenuModifierID),
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod2ID = MenuModifiers.MenuModifierID),
        (SELECT MenuModifiers.MenuModifierText FROM MenuModifiers WHERE OrderTransactions.Mod3ID = MenuModifiers.MenuModifierID)
        FROM (OrderTransactions INNER JOIN MenuItems ON OrderTransactions.MenuItemID = MenuItems.MenuItemID)
        INNER JOIN OrderHeaders ON OrderHeaders.OrderID = OrderTransactions.OrderID
        WHERE MenuItems.MenuItemText = '"+params[:liquor_type]+"'
        AND OrderHeaders.OrderDateTime > #"+params[:start_date].to_s+"#
        AND OrderHeaders.OrderDateTime <= #"+params[:end_date].to_s+"#;")

      @liquor_sales_data = db.data
      @liquor_sales_data.sort!
      @liquor_type = params[:liquor_type]
      @start_date = params[:start_date]
      @end_date = params[:end_date]

      render :partial => "liquor_sales"
    end
  end

  def export_liquor_sales
    if params[:start_date]
      start_date = params[:start_date]
      end_date = params[:end_date]
      liquor_type = params[:liquor_type]

      @outfile = liquor_type+"_"+start_date+"_to_"+end_date+".csv"
      send_data Report.liquor_sales_to_csv(start_date, end_date, liquor_type), :type => 'text/csv; charset=iso-8859-1; header=present', :disposition => "attachment; filename=#{@outfile}"
    end
  end
  
  
  
  
  
  
end#end Report Controller