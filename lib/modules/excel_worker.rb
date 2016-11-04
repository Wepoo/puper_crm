module ExcelWorker
  def make_file (offer)
    deal = offer.deal
    contact = offer.agent.contact
    manager = offer.agent.manager
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet :name => 'Deal'
    
    sheet.row(0).push 'Name'
    sheet.row(0).push offer.agent.name
    sheet.row(1).push 'Company name'
    sheet.row(1).push offer.agent.company_name
    sheet.row(2).push 'Delivery time'
    sheet.row(2).push Time.now + 1.month

    sheet.row(0) << '' << ''
    sheet.row(0).concat %w{Full_name Position Phone Email Skype Fax}
    sheet.row(1) << '' << 'Contact'
    sheet.row(1) << contact.full_name << contact.position << contact.phone << contact.email
    sheet.row(1) << contact.skype << contact.fax

    sheet.row(2) << '' << 'Manager'
    sheet.row(2) << manager.full_name << manager.position << manager.phone << manager.email
    sheet.row(2) << manager.skype << manager.fax

    sheet.row(5).push 'Product name'
    sheet.row(5) << 'SKU' << 'Currency' << 'Promo' << 'Unit price(promo)' << 'Unit price(regular)'
    sheet.row(6).push deal.name
    sheet.row(6) << deal.sku << deal.currency.name << deal.promo << deal.promo_unit_price << deal.unit_price << deal.unit_type
    sheet.row(5).push 'Unit'
    sheet.row(5) << 'Plan Level'
    7.times do |index|
      sheet.merge_cells(6, index, 6 + deal.plans.count, index)
    end
    deal.plans.each_with_index do |plan, index|
      sheet.row(6 + index).insert 7, plan.name
    end
    plan = deal.plans.first
    time = plan.end_at - plan.start_at
    decades = (time/24/3600/10).to_i
    decades.times do |i|
      sheet.row(5) << (plan.start_at + 10.days*i).to_s
    end
    quarters = (time/24/3600/91).to_i
    quarters.times do |i|
      sheet.row(6) << '-'
      sheet.merge_cells(6, 8+8*i, 6, 8+(i+1)*8)
    end
    months = (time/24/3600/30).to_i - 1
    months.times do |m|
      sheet.row(7) << '-'
      sheet.merge_cells(7, 8 + 4*m, 7, 11+4*m) 
    end
    sheet.row(7) << '-'

    sheet.row(5) << 'Total' << 'Due date' << 'Accept' << 'Descr'

    deal.plans.each_with_index do |plan, index|
      sheet.row(6 + index).insert 8+decades, "=SUM(I#{7+index}:Q#{7+index})"
      sheet.row(6 + index).insert 11+decades, "сумма"
    end
    book.write 'tmp/offer.xls'
  end

  def from_file
    book = Spreadsheet.open 'tmp/new_offer.xls'
    sheet = book.worksheet 0
    col = 17
    result = {
      email: sheet[1,7],
      strategic: sheet[6,col].value,
      perspect: sheet[7,col].value,
      operational: sheet[8,col].value,
      current: sheet[9,col].value
    }
  end

end
