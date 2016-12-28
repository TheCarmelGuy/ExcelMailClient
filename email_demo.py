import smtplib
import xlwings as xw
import env



#def gen_range( ):




def smtp_init(server):

    
    #get all the server data that you need to use email client
    server.starttls()
    server.login(env.email_username, env.email_pass);


def main():
  
    #set up server 
    server = smtplib.SMTP('smtp.gmail.com', 587)
    
    smtp_init(server)    

 
    #load proper excel spreadsheet
    wb = xw.Book('emails.xlsx')
    #assume that here A column contins all the emails and B contain the info that is needed
    datasheet = wb.sheets['emails']


    count = 0
    #iterate with customizable message
    for email, amount in zip(datasheet.range('A1:A2').value, datasheet.range('B1:B2').value):
    
        template_msg = str(amount) + " is how much you owe. Thank you!"   
        server.sendmail(env.email_username,str(email), template_msg)
        print("Payment request was sent to " + email + " for the amount of $" + str(amount)) 
        count += 1
        
    server.quit()
    print str(count) + " emails have been proccessed" 
if __name__=="__main__": 
    main()
