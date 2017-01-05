import smtplib
import xlwings as xw
import env
import sysi, getopt


#return array of data points in excel sheet
#requires a lower bound, upper bound, and an excel data sheet as inputs

def gen_range(lower_bound, upper_bound,sheet):

    return sheet.range(lower_bound +':'+ upper_bound).value


def smtp_init(server):

    
    #get all the server data that you need to use email client
    server.starttls()
    server.login(env.email_username, env.email_pass);

def emailDispatcher(filename, beginning, end):
  
    #set up server 
    server = smtplib.SMTP('smtp.gmail.com', 587)
    
    smtp_init(server)    

 
    #load proper excel spreadsheet
    wb = xw.Book(filename)
    #assume that here A column contins all the emails and B contain the info that is needed
    datasheet = wb.sheets['emails']


    count = 0
	
	#the character can be changed to set a particular column in the data set
	#NOTE This is for setting the column of the email account information
	emailBase = 'A' + beginning
	emailBound = 'A' + end
		
	#the character can be changed to set a particular column in the data set
	#NOTE This is for setting the column of the dues data
	duesBases = 'B' + beginning 
	duesBound = 'B' + end	

    #iterate with customizable message
    for email, amount in zip(gen_range(emailBase, emailBound, datasheet), gen_range(duesBase,duesBound, datasheet)):
    
        template_msg = str(amount) + " is how much you owe. Thank you!"   
        server.sendmail(env.email_username,str(email), template_msg)
        print("Payment request was sent to " + email + " for the amount of $" + str(amount)) 
		count += 1
        
    server.quit()
    print str(count) + " emails have been proccessed"
	print("Check Log.txt for the information that has been sent out") 


def main(argv):
	
		







if __name__=="__main__": 
    main()
