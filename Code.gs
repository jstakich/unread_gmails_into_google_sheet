function onOpen()
{
     var ActiveSheet = SpreadsheetApp.getActiveSpreadsheet();
     var Sheet1  = ActiveSheet.getActiveSheet();
     Sheet1.getRange('A1:C40').clearContent();
     var threads = GmailApp.getInboxThreads(0,100);
     var cell = Sheet1.getRange('A1');
     var row = -1;
   
    for (var col = 0; col < threads.length; col++)
            {
            var AllEmailMessages = threads[col].getMessages();
                   for(var TheEmailMsg in AllEmailMessages)
                   {
                     row = row + 1
                     
                     
                     
                     if (threads[col].isUnread()) {
                         cell.offset(row, 0).setValue('=hyperlink("https://mail.google.com/","Go")')
                         cell.offset(row, 1).setValue(AllEmailMessages[TheEmailMsg].getSubject());              
                         cell.offset(row, 2).setValue(AllEmailMessages[TheEmailMsg].getFrom());
       }
       
       
                    
                         
             
                        if (row == 4)
                        {
                       // Browser.msgBox('hello world', Browser.Buttons.OK_CANCEL);
                          return;
                        }
                     
                  
                       
                    
                     
                     
                   }
             }
             row++;
  }

