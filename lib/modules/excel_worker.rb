module Excel_worker
  def make_file
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
    sheet.row(6) << deal.sku << deal.currency.name << deal.promo << deal.promo_unit_price << deal.unit_price
    sheet.row(5).push 'Unit'
    sheet.row(5) << 'Plan Level'
    6.times do |index|
      sheet.merge_cells(6, index-1, 6 + deal.plans.count, index-1)
    end
    deal.plans.each_with_index do |plan, index|
      sheet.row(6 + index).insert 7, plan.name
    end

    book.write 'tmp/offer.xls'
  end

end
