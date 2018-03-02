using System;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using Limilabs.Mail;
using Limilabs.Client.IMAP;
using Limilabs;
using System.Collections.Generic;
using System.Security;
using System.Threading;
using System.Security.Cryptography;


namespace ConsoleSpeech
{
    class ConsoleSpeechProgram
    {
        static SpeechSynthesizer sp = new SpeechSynthesizer(); //new instance of SpeechSynthesizer module referred to as 'sp'
        static SpeechRecognitionEngine sre; //new instance of SpeechRecognitionEngine module referred to as 'sre'
        static bool done = false; //establishes a boolean later used to end the program
        static bool speechOn = true; //establishes boolean that later functions as a speech on/off switch


        static void Main(string[] args)
        {
            try
            {
                Console.Title = "Equal Access Software Environment (EASE) Alpha Build 32"; //writes window title
                sp.SetOutputToDefaultAudioDevice(); //sets output to default audio device
                sp.SelectVoice("Microsoft David Desktop"); //Sets voice, default is 'Microsoft David Desktop'
                if (File.Exists(@"C:\Users\grace\source\repos\SynthesisResponse Alpha 33\SynthesisResponse Alpha 33\Password.txt") && (File.Exists(@"C:\Users\grace\source\repos\SynthesisResponse Alpha 33\SynthesisResponse Alpha 33\EmailAddress.txt"))) //IMPROVE LATER
                {
                    Console.WriteLine("Successfully found passwords");
                }
                else
                {
                    Console.WriteLine("Email address and password not found. Enter your email address now.");
                    sp.SpeakAsync("Email address and password not found. Please enter your email address now.");
                    string email = Console.ReadLine();
                    
                    System.IO.File.WriteAllText(@"C:\Users\grace\source\repos\SynthesisResponse Alpha 33\SynthesisResponse Alpha 33\EmailAddress.txt", email);

                    Console.WriteLine("Thank you. Now, please enter your password.");
                    sp.SpeakAsync("Thank you. Now, please enter your password.");
                    string password = ReadPassword();
                    Console.WriteLine("\nIf nobody's looking over your shoulder");
                    Console.WriteLine("I'm going to confirm the password you gave");
                    Console.WriteLine("After 5 seconds, all input will be cleared and the program will restart.");
                    Console.WriteLine("If you entered the wroing password, close the program and re-open it to enter it again.");
                    Console.WriteLine("\nPress any key to see it now");
                    Console.ReadKey(true);
                    Console.Write("\nThe password entered is : " + password);
                    Thread.Sleep(5000);
                    Console.Clear();
                    System.IO.File.WriteAllText(@"C:\Users\grace\source\repos\SynthesisResponse Alpha 33\SynthesisResponse Alpha 33\Password.txt", password);

                }
                Console.WriteLine("\nWelcome to the Equal Access Software Environment. To start, please say, 'send email' to send an email, or 'read email'  to check inbox. If you want to see the command list again, please say 'inspect entire command list'."); //Writes line
                sp.Speak("Welcome to the Equal Access Software Environment. To start, please say, 'send email' to send an email, or 'read email' to check inbox. If you want to see the command list again, please say 'inspect entire command list'."); //Speaks
                CultureInfo ci = new CultureInfo("en-us"); //sets CultureInfo to that of English - United States to better recognition
                sre = new SpeechRecognitionEngine(ci); //merges previous sre with the cultureinfo
                sre.SetInputToDefaultAudioDevice(); //Sets input to default audio device
                sre.SpeechRecognized += sre_SpeechRecognized; //the module sre.SpeechRecognized = sre.SpeechRecognized + object sre_SpeechRecognized
                Choices ch_StartStopCommands = new Choices(); //makes new Choices list
                ch_StartStopCommands.Add("speech on");
                ch_StartStopCommands.Add("speech off");
                ch_StartStopCommands.Add("finish");
                ch_StartStopCommands.Add("inspect entire command list"); //All of these add commands to the choices list
                ch_StartStopCommands.Add("send email");
                ch_StartStopCommands.Add("read email");
                GrammarBuilder gb_StartStop = new GrammarBuilder(); //Makes new GrammarBuilder
                gb_StartStop.Append(ch_StartStopCommands); //Adds chooices to the GrammarBuilder
                Grammar g_StartStop = new Grammar(gb_StartStop); //Makes a grammar dictionary referencing the system's grammar and the grammarbuilder
                sre.LoadGrammarAsync(g_StartStop); //actually loads the grammar for use
                sre.RecognizeAsync(RecognizeMode.Multiple); //starts recognition, hearing for multiple words
                while (done == false) {; } //checks if the bool done is false, if it is it progresses
                Console.WriteLine("\nHit <enter> to close shell\n"); //writes
                Console.ReadLine(); //waits for input
            }
            catch (Exception ex1) //exception handler, nothing special
            {
                Console.WriteLine(ex1.Message);
                Console.ReadLine();
            }
        } // Main
        static void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) //once the speech is recognized, it references this object
        {
            string smtpassword = File.ReadAllText(@"C:\Users\grace\source\repos\SynthesisResponse Alpha 33\SynthesisResponse Alpha 33\Password.txt");
            string smtemail = File.ReadAllText(@"C:\Users\grace\source\repos\SynthesisResponse Alpha 33\SynthesisResponse Alpha 33\EmailAddress.txt");
            SmtpClient cv = new SmtpClient("smtp.gmail.com", 587)
            {
                EnableSsl = true, //enables ssl encryption
                Credentials = new NetworkCredential(smtemail, smtpassword) //puts in credentials, heres just a sample one
            }; //new smtp client
            string txt = e.Result.Text; //a string with the result args, what it thinks that you said
            float confidence = e.Result.Confidence; //float
            Console.WriteLine("\nRecognized: " + txt); //writes line 'recognized <startstopcommand>'
            if (confidence < 0.60) return; //If the confidence value is less than 60%, it discards it
            if (txt.IndexOf("speech on") >= 0) //if recognizes command, then performs function. Same method used for all commands.
            {
                Console.WriteLine("Speech is now ON");
                speechOn = true;
            }
            if (txt.IndexOf("speech off") >= 0)
            {
                Console.WriteLine("Speech is now OFF");
                speechOn = false;
            }
            if (speechOn == false) return;
            if (txt.IndexOf("finish") >= 0)
            {
                ((SpeechRecognitionEngine)sender).RecognizeAsyncCancel(); //stop recognition
                done = true; //sets 'done' to true
                Console.WriteLine("(Speaking: Thank you for using me.)"); //writes
                sp.Speak("Thank you for using me."); //speaks
            }
            if (txt.IndexOf("inspect entire command list") >= 0)
            {
                Console.WriteLine(@"Here is the entire command list currently:
                    speech on - turns speech functions on
                    speech off - turns speech functions off
                    send email - sends an email to a specified recipient
                    read email - checks your inbox
                    finish - terminates the program
                    inspect entire command list - shows this list");
                sp.SpeakAsync(@"Here is the entire command list currently:
                    speech on - turns speech functions on
                    speech off - turns speech functions off
                    send email - sends an email to a specified recipient
                    read email - checks your inbox
                    finish - terminates the program
                    inspect entire command list - shows this list");
            }

            if (txt.IndexOf("read") >= 0 && txt.IndexOf("email") >= 0)
            {

                using (Imap imap = new Imap())
                {
                    imap.ConnectSSL("smtp.gmail.com"); //connects to smtp server
                    imap.UseBestLogin(smtemail, smtpassword); //logs in

                    imap.SelectInbox();
                    List<long> uids = imap.Search(Flag.Unseen); //searches for emails with the 'unseen' flag set

                    foreach (long uid in uids)
                    {
                        var eml = imap.GetMessageByUID(uid);
                        IMail email = new MailBuilder()
                            .CreateFromEml(eml);

                        Console.WriteLine(email.From); //Shows message on screen
                        Console.WriteLine(email.Subject);
                        Console.WriteLine(email.Text);
                        sp.SpeakAsync(Convert.ToString(email.From)); //Reads message
                        sp.SpeakAsync(Convert.ToString(email.Subject));
                        sp.SpeakAsync(Convert.ToString(email.Text));

                    }
                    imap.Close(); //Finishes
                }
            }
            if (txt.IndexOf("send") >= 0 && txt.IndexOf("email") >= 0)
            {
                try
                {
                    string email = "placeholdr"; //unused
                    string recipient = "placeholder";
                    string subject = "test";
                    string body = "text";
                    if (email == "placeholder")//this entire if statement is unnescessary and will never show up
                    {
                        Console.WriteLine("There is currently no email set. Please enter your email.");
                        sp.SpeakAsync("There is currently no email set. Please enter your email.");
                        email = Console.ReadLine();
                        Console.WriteLine(email);

                    }
                    else
                    {
                        Console.WriteLine("Please enter the recipient's email address.");//writes
                        sp.SpeakAsync("Please enter the recipient's email address.");//speaks
                        recipient = Console.ReadLine();//sets string to user input
                        Console.WriteLine("Please enter the subject of the message.");//same thing
                        sp.SpeakAsync("Please enter the subject of the message.");
                        subject = Console.ReadLine();
                        Console.WriteLine("Please type the body of the message.");
                        sp.SpeakAsync("Please tyype the body of the message.");
                        body = Console.ReadLine();
                        Console.WriteLine("Sending mail...");
                        sp.SpeakAsync("Sending mail");
                        cv.Send(smtemail, recipient, subject, body);//sends message with specified input
                        Console.WriteLine("Email sent succesfully! :)");
                        sp.SpeakAsync("Email sent successfully");

                    }

                }
                catch (Exception ex) //Exception handler
                {
                    Console.WriteLine("The email refused to send :(");
                    sp.SpeakAsync("The email didn't send. The crash log is as follows");
                    Console.WriteLine(ex);   //Should print stacktrace + details of inner exception

                    if (ex.InnerException != null)
                    {
                        Console.WriteLine("InnerException is: {0}", ex.InnerException);
                        Console.ReadKey();
                    }
                }

            }
            // sre_SpeechRecognized

        }

        public static string ReadPassword()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        // remove one character from the list of password characters
                        password = password.Substring(0, password.Length - 1);
                        // get the location of the cursor
                        int pos = Console.CursorLeft;
                        // move the cursor to the left by one character
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        // replace it with space
                        Console.Write(" ");
                        // move the cursor to the left by one character again
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }

            // add a new line because user pressed enter at the end of their password
            Console.WriteLine();
            return password;
        }// Program

    } // ns

}