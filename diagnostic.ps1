#########################################################
############# Connect Wifi and Testing ##################
#########################################################
#create Wi-Fi xml profile
#netsh wlan export profile "DNCLtechzone" key=clear folder=D:\

#import the network profile
netsh wlan add profile filename="D:\Wi-Fi-DNCLtechzone.xml" Interface="Wi-Fi" user=current
netsh wlan add profile filename="E:\Wi-Fi-DNCLtechzone.xml" Interface="Wi-Fi" user=current
#connect WiFi
netsh wlan connect ssid= "DNCLtechzone" name= "DNCLtechzone" 

# ---------------------------------------------------------------------------------
#This is where all the test will take place

#open camera
start microsoft.windows.camera:

#open testing websites
$navOpenInBackgroundTab = 0x1000;
$ie = new-object -com InternetExplorer.Application
$ie.Navigate2("https://www.youtube.com/watch?v=2ZrWHtvSog4");
$ie.Navigate2("http://deadpixelbuddy.com", $navOpenInBackgroundTab);
$ie.Navigate2("http://en.key-test.ru", $navOpenInBackgroundTab);
$ie.Visible = $true;
devmgmt.msc

