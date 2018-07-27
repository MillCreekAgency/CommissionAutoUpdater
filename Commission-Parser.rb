#!/usr/local/bin/env ruby

require 'rubyXL'
require 'selenium-webdriver'


class CommissionParser

  
  def initialize
    workbook = RubyXL::Parser.parse("MILL_CREEK.xlsx")
    @worksheet = workbook.worksheets[0] 
    @start_index = self.find_start(@worksheet)
  end
  

  def find_start(worksheet)
    i = 0
    worksheet.each do |row|
      
      if row != nil and row[0] != nil and row[0].value == 'COMPANY'
        return i + 2
      end
      i += 1
    end
    return -1
  end

  def display_policies
    @policies.each do |policy|
      policy.each do |key, value|
        puts "#{key}: #{value}"
      end
      puts ""
    end
  end

  def print_pre pre
    pre.each do |key, value|
        puts "#{key}: #{value}"
    end
  end


  def find_insureds
    @policies = Array.new
    previous_insured = {}
    (@start_index...@worksheet.sheet_data.size).each do |row|
      row_sheet = @worksheet[row]
      if self.is_end?(row_sheet)
        @policies.push(previous_insured)
        break
      end
      if same_policy?(row_sheet, previous_insured)
        # puts previous_insured[:commission]
        #puts self.get_commission_value(row_sheet[7])
        previous_insured[:commission] += self.get_commission_value(row_sheet[7])
      else
        @policies.push(previous_insured)
        previous_insured = {
          policy_num: self.get_cell_value(row_sheet[2]),
          name: self.get_cell_value(row_sheet[1]),
          date: self.get_cell_value(row_sheet[3]),
          company: self.get_cell_value(row_sheet[0]),
          commission: self.get_commission_value(row_sheet[7])
        }
      end
    end
  end

  def get_commission_value(cell)
    if cell == nil or cell.value == nil
      return 0
    else
      return cell.value
    end
  end

  def get_cell_value(row)
    return row == nil ? nil : row.value
  end

  def same_policy?(row, pre_ins)
    name = get_cell_value(row[1])
    policy_num = get_cell_value(row[2])
    date = get_cell_value(row[3])
    same_name = ((name == nil) or (name == pre_ins[:name])) 
    same_policy = ((policy_num == nil) or (policy_num == pre_ins[:policy_num]) )
    same_date = ((date == nil) or (date == pre_ins[:date]))
    (same_name and same_policy) and same_date
  end

  def is_end?(row)
    if (row != nil)
      return self.all_nil(row)
    end
      return true
  end

  def all_nil(row)
    row.cells.each do |cell|
      if cell != nil and cell.value != nil
        return false
      end
    end
    return true
  end
end


class WebDriver

  def initalize
    @driver = Selenium::WebDriver.for :chrome
    @driver.navigate.to 'https://app.qqcatalyst.com/Contacts/MGA/Details/7152'
    print 'Please enter password:'
	pass = STDIN.noecho(&:gets).chomp
    puts ""

    if @driver.current_url.include? 'login.qqcatalyst.com'
      # Logs in
      emailField = driver.find_element(name: 'txtEmail')
      emailField.send_keys 'tory@millcreekagency.com'
      passField = driver.find_element(id: 'txtPassword')
      passField.send_keys pass
      @driver.find_element(id: 'lnkSubmit').click

      sleep(1)

      if @driver.current_url.include? 'login.qqcatalyst.com'
        yes_button = driver.find_element(id: 'lnkCancel')
        yes_button.click
      end
    end

    # Go to reconcile
    reconcile_button = @driver.find_elements(class: "ReconileContact")[0]
    reconcile_button.click

    until @driver.current_url.include? 'ReconcileWorkflow?'
      sleep(1)
    end

    @driver.find_element(id: "CarrierMGA_Paying_Agency_Commissions").click

    until @driver.current_url.include? 'ReconcileSelectPoilices'
      sleep(1)
    end
  end

  def update_policy policy_num, name, commission, date, company 
    if policy_num != nil
      @driver.find_element(name: "sch-policyNumber").send_keys policy_num
    else
      @driver.find_element(name: "sch-insured").send_keys name
    end
    
    end
  end


end

cp = CommissionParser.new
cp.find_insureds
cp.display_policies
