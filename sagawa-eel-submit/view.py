import eel
import desktop
import sagawa

app_name="html"
end_point="index.html"
size=(2000,1300)

@ eel.expose
def main(xlname):
    sagawa.main(xlname)
    
desktop.start(app_name,end_point,size)
# # desktop.start(size=size,appName=app_name,endPoint=end_point)