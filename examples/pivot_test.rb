class RandomReportGenerator
  def date
    Date.today.strftime("%m/%d/%Y")
  end
  def member_id
    @i ||= 0
    @i += 1
  end
  def name
    "John S."
  end
  def gender
    ["Male", "Female"].sample
  end
  def age
    rand(100)
  end
  def city
    ["New York", "Mountain View", "Newark", "Phoenix"].sample
  end
  def state
    ["NY", "CA", "NJ", "AZ"].sample
  end
  def parenting
    "Foo"
  end
  def student
    "Bar"
  end
  def income
    "Bar"
  end
  def education
    "Bar"
  end
  def answer
    ["Yes", "No", "Maybe", "I dont know"].sample
  end
  def run
    package  = Axlsx::Package.new
    workbook = package.workbook
 
    workbook.add_worksheet(:name => "Data Sheet") do |sheet|
      sheet.add_row [
        "Date", "Member ID", "Name", "Gender", "Age", "City", "State",
        "Parenting Status", "Student Status", "Income", "Education", "Answer"
      ]
      30.times do
        sheet.add_row [date, member_id, name, gender, age, city, state,
                       parenting, student, income, education, answer]
      end
    end
 
    workbook.add_worksheet(:name => "Summary") do |sheet|
      pivot_table = Axlsx::PivotTable.new 'A1:B15', "A1:L31", workbook.worksheets[0]
      pivot_table.rows = ['Answer']
      pivot_table.data = [{:ref => "Member ID", :subtotal => "count"}]
      sheet.pivot_tables << pivot_table
    end
 
    package.serialize("pivot_table.xlsx")
  end
end
