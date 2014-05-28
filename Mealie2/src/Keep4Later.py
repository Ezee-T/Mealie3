import smtplib
server = smtplib.SMTP('smtp.gmail.com', 587)
 
#Next, log in to the server
server.login("vbezee@gmail.com", "realmad23")
 
#Send the mail
msg = "\nHello!" # The /n separates the message from the headers
server.sendmail("vbezee@gmail.com", "Thubelihle.Zondi@hulamin.co.za", msg)