#########################################################
############# Connect Wifi and Testing ##################
#########################################################
#create Wi-Fi xml profile
#netsh wlan export profile "NETGEAR61" key=clear folder=E:\

#import the network profile
netsh wlan add profile filename="D:\Wi-Fi-NETGEAR61.xml" Interface="Wi-Fi" user=current
netsh wlan add profile filename="E:\Wi-Fi-NETGEAR61.xml" Interface="Wi-Fi" user=current
#connect WiFi
netsh wlan connect ssid= "NETGEAR61" name= "NETGEAR61" 

# ---------------------------------------------------------------------------------
#This is where all the test will take place

#open camera
#start microsoft.windows.camera:

#open testing websites
$navOpenInBackgroundTab = 0x1000;
$ie = new-object -com InternetExplorer.Application
$ie.Navigate2("https://www.youtube.com/watch?v=2ZrWHtvSog4");
$ie.Navigate2("http://deadpixelbuddy.com", $navOpenInBackgroundTab);
$ie.Navigate2("http://en.key-test.ru", $navOpenInBackgroundTab);
#$ie.Navigate2("https://www.onlinemictest.com/webcam-test/", $navOpenInBackgroundTab);
$ie.Visible = $true;
devmgmt.msc
