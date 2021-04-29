from win32com.client import Dispatch
import webbrowser
import os
import datetime
print ("Hi , myself NIRU , your personal assistant, I will work on your commands that you enter below , Before that ")
xy =(input("Enter your Name :"))
speak = Dispatch("SAPI.SpVoice")

speak.Speak("welcome")
speak.Speak(xy)
speak.Speak("I am niru and nice to meet you")
speak.Speak ("Let's continue")
print ('Welcome',xy)
print ("General Instruction , you have to give the command by adding the name 'niru' , for example: niru add , niru multiply , niru divide , niru subtract etc .")
print ("to know the different command that you can give , copy this link and paste it in google: ")
print (xy ,"let's start , pls enter the command below :" )

while True :
    
      cmd = input("your command :")

      if cmd == 'niru add':

          x = int(input("Enter a number:"))
          y = int(input("Enter the second number :"))
          z = x + y
          print (z)
          print("answer is :",z)
          speak = Dispatch("SAPI.SpVoice")

          speak.Speak(z)

      elif cmd == 'niru multiply':
          x = int(input("Enter a number:"))
          y = int(input("Enter the second number :"))
          z = x * y
          print (z)
          print("answer is :",z)
          speak = Dispatch("SAPI.SpVoice")

          speak.Speak(z)

      elif cmd == 'niru subtract':
          x = int(input("Enter a number:"))
          y = int(input("Enter the second number :"))
          z = x - y
          print (z)
          print("answer is :",z)
          speak = Dispatch("SAPI.SpVoice")

          speak.Speak(z)

      elif cmd == 'niru divide':
          x = int(input("Enter a number:"))
          y = int(input("Enter the second number :"))
          z = x/y
          print (z)
          print("answer is :",z)
          speak = Dispatch("SAPI.SpVoice")

          speak.Speak(z)

      elif cmd == 'niru add 3':
          x = int(input("Enter a number:"))
          y = int(input("Enter the second number :"))
          a = int(input("Enter the third number :"))
          z = x+y+a
          print("answer is :",z)
          speak = Dispatch("SAPI.SpVoice")

          speak.Speak(z)

      elif cmd == 'niru quit':
            print("Thank you for using Niru",xy)
            print("Click the OK to exit")
            speak = Dispatch("SAPI.SpVoice")

            speak.Speak("Thank you for using Niru")
            speak.Speak("let's meet afterwards")
            speak.Speak("click 'OK' to exit")
            exit()

      elif cmd == 'niru good morning':
            print ("A very good morning",xy)
            speak = Dispatch("SAPI.SpVoice")

            speak.Speak("A very good morning")
            speak.Speak("and Have a nice day")

      elif cmd == 'niru thanks':
            print ("Your welcome",xy)
            speak = Dispatch("SAPI.SpVoice")

            speak.Speak("Your welcome")
            speak.Speak("No thanks between friends")

      elif cmd == 'niru website':
            print ("Just a second")
            speak = Dispatch("SAPI.SpVoice")

            speak.Speak("sure , opening via MS edge")
            webbrowser.open('https://rzmyf5fannpwfsdvqpwq5g-on.drv.tw/niru123/NIRUNEW.html', new=2)
            
      elif cmd == 'niru shutdown':
            speak.Speak("are you sure to shutdown your system")
            shutdown = input("Shutdown (yes / no): ")
            
	       
            speak = Dispatch("SAPI.SpVoice")
            
            

            if shutdown == 'no':
	            speak.Speak("Ok")
            else:
	            os.system("shutdown /s /t 1")
	            speak.Speak("Shuting down the system")

      elif cmd == 'niru time':
            x = datetime.datetime.now()
            print(x)
            
           
	            

            


      else:
         print("Please enter a valid input")
         speak = Dispatch("SAPI.SpVoice")

         speak.Speak("sorry I don't know that one")
          

      

          



