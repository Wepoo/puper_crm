class OffersMailer < ApplicationMailer
  include Excel_worker

  def distribute(offer)
    Excel_worker.make_file (offer)
    contact_email = offer.agent.contact.email
    attachments['offer.xls'] = { mime_type: 'application/xls', content: File.read('tmp/offer.xls') }
    subject = offer.type
    p '======================================'
    p contact_email.to_s
    p subject.to_s
    p '======================================'
    mail to: contact_email, subject: subject
  end

end
