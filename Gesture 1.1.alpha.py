
import cv2
import numpy as np
import win32com.client
import time

def show_webcam(mirror=False):
    #app = win32com.client.Dispatch("PowerPoint.Application")
    #presentation = app.Presentations.Open(FileName=u'C:\\Users\\Aman Jain\\Downloads\\Lecture 7.ppt', ReadOnly=1)
    #presentation.SlideShowSettings.Run()

    cam = cv2.VideoCapture(0)
    cam.set(3,640)
    cam.set(4,480)
    while True:
        ret_val, img = cam.read()
        img_hsv=cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        if mirror: 
            img = cv2.flip(img, 1)
            img_hsv=cv2.cvtColor(img, cv2.COLOR_BGR2HSV)



        # lower mask (0-10)
        lower_red = np.array([0,90,30])
        upper_red = np.array([10,255,255])
        mask0 = cv2.inRange(img_hsv, lower_red, upper_red)

        # upper mask (170-180)
        lower_red = np.array([170,90,30])
        upper_red = np.array([180,255,255])
        mask1 = cv2.inRange(img_hsv, lower_red, upper_red)

        # join my masks
        mask = mask0+mask1

        # set my output img to zero everywhere except my mask
        output_img = img.copy()
        output_img[np.where(mask==0)] = 0
        median = cv2.bilateralFilter(output_img,9,75,75)

        #Noise removal technique.. quite slow !
        #dst = cv2.fastNlMeansDenoisingColored(output_img,None,10,10,7,21)
        
        count = 0
        for i in range (50,351):
            if(output_img[i][100].all() != 0):
                count=count+1;
        if (count>20):
            cv2.line(output_img, (100,50), (100,350), (255,0,0), thickness=2, lineType=8, shift=0)
            #presentation.SlideShowWindow.View.Next()
        else:
            cv2.line(output_img, (100,50), (100,350), (0,255,0), thickness=2, lineType=8, shift=0)

        cv2.circle(img, (320,240), 50, (0,255,0), thickness=2, lineType=8, shift=0)

        ret,thresh1 = cv2.threshold(output_img,127,255,cv2.THRESH_BINARY)
        image, contours, hierarchy = cv2.findContours(thresh1,cv2.RETR_CCOMP,cv2.CHAIN_APPROX_SIMPLE)
        img = cv2.drawContours(thresh1, contours, -1, (0,255,0), 3)
        cv2.imshow('my webcam', img)
        t = img[240][320]
        print t
        if cv2.waitKey(1) == 27: 
            break  # esc to quit
    cv2.destroyAllWindows()

def main():
    show_webcam(mirror=True)

if __name__ == '__main__':
    main()