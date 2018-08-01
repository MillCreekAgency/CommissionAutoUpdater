#!/usr/local/bin/env ruby
require 'io/console'
require 'rubyXL'
require 'selenium-webdriver'


USER_EMAIL='tory@millcreekagency.com'

class CommissionParser

  def initialize
    puts "Enter filepath: "
    filepath = gets.chomp.strip
    workbook = RubyXL::Parser.parse(filepath)
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

  def display_policies policies
    results = ""
    policies.each do |policy|
      policy.each do |key, value|
        if value.class.name == "DateTime"
          str = "#{key}: #{value.strftime("%m/%d/%Y")}"
        else
          str ="#{key}: #{value}"
        end
        puts str
        results += str + "\n"
      end
      puts ""
      results += "\n"
    end
    out_file = File.new(File.expand_path("~/Desktop/ToBeDone-Commission.txt"), "w")
    out_file.puts results
    out_file.close
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
        if previous_insured != {}
          @policies.push(previous_insured)
        end
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

  def update
    web_driver = WebDriver.new

    not_found = []
    @policies.each do |policy|
      if not web_driver.update_policy policy
        not_found.push(policy)
      end
    end
    # display_policies not_found
    done = false

    until done == true do
      puts "Finished? [Y/n]"
      answer = gets.chomp
      if not answer.include? 'n'
        done = true
      end
    end
  end
end


class WebDriver

  def initialize
    @driver = Selenium::WebDriver.for :chrome
    @driver.navigate.to 'https://app.qqcatalyst.com/Contacts/MGA/Details/7152'
    
    print 'Please enter password:'
	pass = STDIN.noecho(&:gets).chomp
    puts ""

    if @driver.current_url.include? 'login.qqcatalyst.com'
      # Logs in
      emailField = @driver.find_element(name: 'txtEmail')
      emailField.send_keys USER_EMAIL
      passField = @driver.find_element(id: 'txtPassword')
      passField.send_keys pass
      @driver.find_element(id: 'lnkSubmit').click

      sleep(1)

      if @driver.current_url.include? 'login.qqcatalyst.com'
        yes_button = @driver.find_element(id: 'lnkCancel')
        yes_button.click
      end
    end

    sleep(1)

    # Go to reconcile
    reconcile_button = @driver.find_elements(class: "ReconcileContact")[0]
    reconcile_button.click


    until @driver.current_url.include? 'ReconcileWorkFlow'
      sleep(1)
    end

    sleep(1)
    @driver.find_element(id: "CarrierMGA_Paying_Agency_Commissions").click
    
    # Click next button
    buttons = @driver.find_elements(class: "basic-page-next")
    buttons.each do |button|
      if (button.attribute("innerHTML").include? "Next") && button.displayed?
        button.click
      end
    end

    until @driver.current_url.include? 'ReconcileWorkflow'
      sleep(1)
    end
  end

  def update_policy policy 
    sleep(1)
    search_policy = @driver.find_element(name: "sch-policyNumber")
    scroll_to search_policy
    search_policy.clear
    if policy[:policy_num] != nil
      search_policy.send_keys policy[:policy_num].tr(' ', '')
    else
      return false 
    end

    button = find_policy_button @driver.find_elements(class: "ResetTutorialsButton")
    button.click

    sleep(2)
    
    if @driver.find_elements(class: "simplemodal-wrap")[0] != nil
      buttons = @driver.find_elements(class: "close")
      buttons.each do |close_button|
        if (close_button.attribute("innerHTML").include? "OK") && close_button.displayed?
          close_button.click
        end
      end
      return false
    end


    # Get search results
    search_results = @driver.find_elements(class: "row-one")
    search_results.each do |result|
      if match_result result, policy[:policy_num], policy[:date]
        addCommission result, policy[:commission]
        break
      end
    end
  end

  def addCommission result, amount
    balance = result.find_element(class: "balance-amt")
    scroll_to balance
    balance.click
    input = result.find_element(class: "currencyVal")
    scroll_to input
    input.clear
    input.send_key(amount.to_s)
  end



  def match_result result, policy_num, date
    date_title = false
    policy_num_title = false
    result.find_elements(tag_name: "td").each do |cell|
      title = cell.property("title")
      if (title.include? policy_num)
        policy_num_title = true
      end
      if (title.include? format_date(date))
        date_title = true
      end
    end
    if date_title and policy_num_title
      return true
    end
    return false
  end

  def format_date date
    if date.class.name == "DateTime" || date.class.name == "Date" 
      return date.strftime("%m/%d/%Y")
    else 
      if date.index("/") == 1
        date = '0' + date
      end

      if date.index('/', 3) == 4
        date = date.insert(3, '0')
      end
      return date
    end
  end

  def find_policy_button elements
    elements.each do |element|
      if (element.attribute("innerHTML").include? "Find Policies") && element.displayed?
        return element
      end
    end
  end

  def scroll_to element
    #@driver.action.move_to(element).perform
    #@driver.execute_script("arguments[0].scrollIntoView(true);", element)
    @driver.execute_script("window.scrollTo(arguments[0], arguments[1]);", element.location.x, element.location.y - 200)
    sleep(1)
  end

  


end

cp = CommissionParser.new
cp.find_insureds
cp.update
