#########################################################
############# Connect Wifi and Testing ##################
#########################################################

#create Wi-Fi xml profile
netsh wlan export profile "wifi_name_in_network" key=clear folder=D:\

#import the network profile
netsh wlan add profile filename="D:\wifi_file.xml" Interface="Wi-Fi" user=current

#connect WiFi
netsh wlan connect ssid= "wifi_name" name= "wifi_name"
# ---------------------------------------------------------------------------------
#This is where all the test will take place

#open camera
#start microsoft.windows.camera:

#open testing websites
$navOpenInBackgroundTab = 0x1000;
$ie = new-object -com InternetExplorer.Application
# speaker/internet connection test
$ie.Navigate2("https://www.youtube.com/watch?v=2ZrWHtvSog4");
# dead/burnt pixel/touchscreen test
$ie.Navigate2("http://deadpixelbuddy.com", $navOpenInBackgroundTab);
# keyboard test
$ie.Navigate2("http://en.key-test.ru", $navOpenInBackgroundTab);
$ie.Visible = $true;
devmgmt.msc
