#!/usr/bin/env python
import xlwt
import rospy
import sys
#to know msg type use "rostopic echo <topic_name>
from vn_100.msg import sensor_data 
from vn_100.msg import ins_data#from vn_100.msg will be for all msgs defined in package
from std_msgs.msg import UInt16,Float64

wb = xlwt.Workbook()
ws1 = wb.add_sheet("SENSOR DATA")
ws2 = wb.add_sheet("INS DATA")
ws3 = wb.add_sheet("DEPTH VALUE")
ws4 = wb.add_sheet("FRONTPITCHSPEED")

count = 0

def heading1():
    global ws1,wb
    style_string = "font: bold on"
    style = xlwt.easyxf(style_string)
    DATA = ["Acceleration[x]","Acceleration[y]","Acceleration[z]","Gyro[x]","Gyro[y]","Gyro[z]","Temp","Pressure","Timestamp"]
    for i in range(0,len(DATA)):
        ws1.write(0,i,DATA[i],style=style)
        ws1.col(i).width=3500
        
def sensor_data_callback(msg):
    global count,ws1,wb
    count = count+1
    global ws1,wb
    DATA = [str(msg.Accel.x),str(msg.Accel.y),str(msg.Accel.z),str(msg.Gyro.x),str(msg.Gyro.y),str(msg.Gyro.z),str(msg.Temp),str(msg.Pressure),str(msg.header.stamp)]
    for i in range(0,9):
        ws1.write(count,i,DATA[i])
        rospy.loginfo("sensor appended,count=")
        rospy.loginfo(count)
        if(i == 8):
            ws1.col(i).width=4700
        else:
            ws1.col(i).width=3500
    wb.save("bag_file_data.xls")

def heading2():
    global ws2,wb
    style_string = "font: bold on"
    style = xlwt.easyxf(style_string)
    DATA = ["Timestamp","YPR[x]","YPR[y]","YPR[z]","quat_data"]
    for i in range(0,len(DATA)):
        ws2.write(0,i,DATA[i],style=style)
        ws2.col(i).width=3500

def ins_data_callback(msg):
    global count,ws,wb
    count = count+1
    global ws2,wb
    DATA = [str(msg.header.stamp),str(msg.YPR.x),str(msg.YPR.y),str(msg.YPR.z),str(msg.quat_data)]
    for i in range(0,len(DATA)):
        ws2.write(count,i,DATA[i])
        rospy.loginfo("ins appended,count=")
        rospy.loginfo(count)
        ws2.col(i).width=3500
    wb.save("bag_file_data.xls")

def heading3():
    global ws3,wb
    style_string = "font: bold on"
    style = xlwt.easyxf(style_string)
    DATA = ["Depth"]
    ws3.write(0,0,DATA[0],style=style)
    ws3.col(0).width=3500

def depth_callback(msg):
    global count,ws3,wb
    count = count+1
    rospy.loginfo("count=")
    rospy.loginfo(count)
    global ws3,wb
    ws3.write(count,0,str(msg.data))
    rospy.loginfo("depth appended,count=")
    rospy.loginfo(count)
    ws3.col(0).width=3500
    wb.save("bag_file_data.xls")

def heading4():
    global ws4,wb
    style_string = "font: bold on"
    style = xlwt.easyxf(style_string)
    DATA = ["FrontPitchSpeed"]
    ws4.write(0,0,DATA[0],style=style)
    ws4.col(0).width=3500

def frontpitchspeed_callback(msg):
    global count,ws4,wb
    count = count+1 #count value should remain updated
    global ws4,wb
    ws4.write(count,1,str(msg.data))
    rospy.loginfo("fps appended,count=")
    rospy.loginfo(count)
    ws4.col(1).width=3500
    wb.save("bag_file_data.xls")

#problem:spacing problem in spreadsheet if subscribers run as separate threads n not sequentially    
if __name__ == "__main__":
    global count
    rospy.init_node("bag_to_excel")
    rospy.loginfo("Working")
    heading1()
    rospy.Subscriber("/vn_100/sensor_data",sensor_data,sensor_data_callback)
    heading2()
    rospy.Subscriber("/vn_100/ins_data",ins_data,ins_data_callback)
    heading3()
    rospy.Subscriber("/depth_value",Float64,depth_callback)
    heading4()
    rospy.Subscriber("/frontpitchspeed",UInt16,frontpitchspeed_callback)
    wb.save("bag_file_data.xls")
    rospy.spin()
