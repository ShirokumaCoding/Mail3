using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using MailKit.Security;
using MimeKit.Utils;

namespace Mail3
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string emailAddress; // your email here
            string password; // your passwork here
            string name; // your email name here
            string destinationAddress; // to whom you are sending it to
            string imageLocation; // complete path to the file;


            // Create a new MimeMessage object
            MimeMessage message = new MimeMessage();

            // Add the sender that will appear in the message
            message.From.Add(new MailboxAddress(name, emailAddress));

            // Add the receiver
            message.To.Add(MailboxAddress.Parse(destinationAddress));

            // Add a message subject
            message.Subject = "I will be damned, this is working.";

            //#region EmailWithAttachment
            //var body = new TextPart("html")
            //{
            //    Text = "<p>I have been trying to reach you about your car's extended warrenty.</p><p><span style=\"Color: #9d682c\">Guess who can send email via code????</span></p>"
            //};

            //// Testing to see if I can attach an image
            //var attachment = new MimePart("image", "jpg")
            //{
            //    Content = new MimeContent(File.OpenRead(imageLocation), ContentEncoding.Default),
            //    ContentDisposition = new ContentDisposition(ContentDisposition.Inline),
            //    ContentTransferEncoding = ContentEncoding.Base64,
            //    FileName = "Bloom_County.jpg"
            //};


            //// Now that we have is created, we need a Multipart container to send both the text and image
            //var multipart = new Multipart("mixed");
            //multipart.Add(body);
            //multipart.Add(attachment);

            //// Finally, compose the message body
            //message.Body = multipart;

            //#endregion

            #region EmailWithInlineImage

            var builder = new BodyBuilder();

            // Need to create the plain text version first?
            builder.TextBody = @"Hey Opus,

I was hoping to borrow your car for my roadtrip back to 1983.

Let me know if it is avaiable.

Bill the Cat.";

            // I think you need to add the image location to some kind of resource pool/bundle so it can be found.
            var image = builder.LinkedResources.Add(imageLocation);

            // No clue what this next part does...Generates a message identifier???
            image.ContentId = MimeUtils.GenerateMessageId();

            // Create an HTML version of the text
            builder.HtmlBody = string.Format(@"<p>Hey Opus,<br /></p>
<br /><p>
I was hoping to borrow your car for my road trip back to 1983.<br /></p>
<br /><p>
Let me know if it is available.<br /></p>
<br /><p>
Bill the Cat<br /></p><br /><br />
<center><img src=""cid:{0}""</center>", image.ContentId);
            message.Body = builder.ToMessageBody();




            #endregion

            // Create a new SMTP client
            SmtpClient smtpClient = new SmtpClient();

            // Create a connection

            try
            {
                smtpClient.Connect("smtp-mail.outlook.com", 587, SecureSocketOptions.StartTls);
                
                // Need to try to authenicate
                smtpClient.Authenticate(emailAddress,password);

                smtpClient.Send(message);

                // This is just to show that the message was sent.
                Console.WriteLine("Message was sent!");
                Console.WriteLine(Environment.CurrentDirectory);
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            finally
            {
                // This code will run regardless of the result of try catch, so its a good place to close the connection
                smtpClient.Disconnect(true);

                // Apparently you need to dipose of the client as well
                smtpClient.Dispose();
            }
        }
    }
}