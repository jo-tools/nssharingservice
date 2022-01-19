# NSSharingService
Xojo example project

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## Description
An example Xojo project to show how to use [NSSharingService](https://developer.apple.com/documentation/appkit/nssharingservice?language=objc) in Xojo-built Applications *(macOS, 64Bit only)*.
The following services are implemented in the example project:
- Compose Email *(Subject, Content, Attachments)*
- Compose Message *(Content/Text, File/Attachments)*
- Send via AirDrop *(Files/Attachments)*

### ScreenShots
Example application - SharingService: **AirDrop**  
![ScreenShot: Airdrop](screenshots/nssharingservice-airdrop.png?raw=true)

Example application - SharingService: **Messages**  
![ScreenShot: Messages](screenshots/nssharingservice-messages.png?raw=true)

Example application - SharingService: **E-Mail**  
![ScreenShot: E-Mail](screenshots/nssharingservice-email.png?raw=true)

## Xojo
### Requirements
[Xojo](https://www.xojo.com/) is a rapid application development for Desktop, Web, Mobile & Raspberry Pi.  

The Desktop application Xojo example project ```NSSharingService.xojo_project``` is using:
- Xojo 2018r4
- API 1

### How to use in your own Xojo project?
1. copy the Module ```modNSSharingService``` to your project.
2. you can then use the provided convenience methods:
   1. ```modNSSharingService.SendViaAirDrop()```
   2. ```modNSSharingService.ComposeMessage()```
   3. ```modNSSharingService.ComposeEmail()```

Note:  
These methods return ```true``` if the sharing service can be invoked, ```false``` if the service is not available.  
The Sharing Service will run asynchrously *(e.g. when not called to show modally within a window)*. That's why you can pass a ```ResultCallbackDelegate```, which will be invoked and reported back with the result later (e.g.: successfully shared, cancelled by user, ...). 

## About
Juerg Otter is a long term user of Xojo and working for [CM Informatik AG](https://cmiag.ch/). Their Application [CMI LehrerOffice](https://cmi-bildung.ch/) is a Xojo Design Award Winner 2018. In his leisure time Juerg provides some [bits and pieces for Xojo Developers](https://www.jo-tools.ch/).

### Contact
[![E-Mail](https://img.shields.io/static/v1?style=social&label=E-Mail&message=xojo@jo-tools.ch)](mailto:xojo@jo-tools.ch)
&emsp;&emsp;
[![Follow on Facebook](https://img.shields.io/static/v1?style=social&logo=facebook&label=Facebook&message=juerg.otter)](https://www.facebook.com/juerg.otter)
&emsp;&emsp;
[![Follow on Twitter](https://img.shields.io/twitter/follow/juergotter?style=social)](https://twitter.com/juergotter)

### Donation
Do you like this project? Does it help you? Has it saved you time and money?  
You're welcome - it's free... If you want to say thanks I'd appreciate a [message](mailto:xojo@jo-tools.ch) or a small [donation via PayPal](https://paypal.me/jotools).  

[![PayPal Dontation to jotools](https://img.shields.io/static/v1?style=social&logo=paypal&label=PayPal&message=jotools)](https://paypal.me/jotools)
