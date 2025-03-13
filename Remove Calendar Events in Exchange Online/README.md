Remove Calendar Events in Exchange Online  
   
# Description  
This PowerShell script removes calendar events in Exchange Online for a specific user within a specified time frame.  
   
# Prerequisites  
- PowerShell installed.  
- Exchange Online module installed.  
- Permissions to manage calendar events in Exchange Online.  
   
# Output  
The script will remove calendar events for the specified user within the defined query window of 360 days. It will cancel organized meetings for the user during this period. # Cancelling these meetings removes them from the user and resource calendars.
   
# Notes  
- Ensure you have the necessary permissions to perform calendar event removal.  
- Replace "EMAIL_PLACEHOLDER" with the actual email address of the user whose calendar events you want to remove.  
- The script will cancel organized meetings for the specified user within the 360-day query window.  
   
# License  
This project is licensed under the MIT License.
