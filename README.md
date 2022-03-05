# Save Screenshot To OneDrive

Save screenshots periodically while the user is logged in.

# How to Use

- Access to https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps and register a new application.
  - Set any URI to the redirect URI (it doesn't use in the process). The platform should be `Web`.
  - Add `Files.ReadWrite.All` and `Files.Read.All` privilege to the application.
  - Create a new client secret.
- Edit `SaveScreenshotToOneDrive.ps1` and set the parameters to access to OneDrive API.
  - Set `Application (Client) ID` to `$ClientId`.
  - Set the redirect URI to `$RedirectURI`.
  - Set the FQDN of OneDrive to `$ResourceId`.
  - Set the client secret to `$AppKey`.
- Run `SaveScreenshotToOneDrive.ps1`.

# Run from task scheduler

- Add a task to launch the content of `SaveScreenshotToOneDrive.ps1`
  - Import the `SaveScreenshotToOneDrive.xml` to the task scheduler.
- Edit the preference of the new task.
- Copy the whole content of `SaveScreenshotToOneDrive.ps1` to the description of the new task.

The task defined in `SaveScreenshotToOneDrive.xml` reads and runs the script from the description of the task itself.
