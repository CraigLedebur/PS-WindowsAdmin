########################################################

     Folder Monitor Script/Windows Service
     Written entrely from scratch by Craig Ledebur
     First commit to GitHub on 16 January 2017
     
########################################################

SUMMARY
---------------
The Windows service, once installed, monitors a hardcoded folder specified
in the C# code. (This will be changed eventually so that it reads from the
registry). When a file is created, in the folder to be monitored, it will
execute a PowerShell script (the path of which is also hard-coded in the
program, and will change eventually).

The script takes over and waits for the file upload to complete, and then
emails the file to the specified address(es) as an attachment. Great for
daily/weekly/monthly reports, automating the delivery of a file via email.

The code was designed to be very flexible and extensible. This was
originally created for a project for a client in which they needed audio
files and transcripts automatically delivered to the recipients as they were
uploaded to a FTP server. The original script ended up being very large, 
well over a thousand lines since it had to do some heavy-duty processing
of the files.

This is a "lite" version, with the project-specific code stripped out. It
works well, and is heavily commented so that you can extend it however you
wish.

INSTALLATION
---------------
After making the proper change to the code for both C# and PowerShell, build
the project and it will generate a Windows Service.

Place the service file in a folder of your choosing, and bring up an
administrator command prompt. Enter the following command:

c:\windows\microsoft.net\framework\v4.0.30319\installutil.exe [path to service exe]

The service will be installed. Go to services.msc to check it and make sure it's
set to Automatic, and to restart in case of failure.

Once the service is running, it will continually monitor the specified folder
(and subfolders) for any new files, and then execute the script. It has been proven
to be 100% reliable on Windows 2012r2, but YMMV (Your Mileage May Vary).

/CL - 16/1/2017
