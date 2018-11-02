import cv2
import numpy as np
import win32com.client
import time


cam = cv2.VideoCapture(0)
cam.set(3,640)
cam.set(4,480)

def show_webcam(mirror=False):
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
	return output_img

def main():
    app = win32com.client.Dispatch("PowerPoint.Application")
    presentation = app.Presentations.Open(FileName=u'C:\\Users\\Aman Jain\\Downloads\\1.pptx', ReadOnly=1)
    presentation.SlideShowSettings.Run()
    while True:
        output_img = show_webcam(mirror=True)

        #Noise removal technique.. quite slow though effective!
        #dst = cv2.fastNlMeansDenoisingColored(output_img,None,10,10,7,21)

        #Forward movement of hand
        count=0
        for i in range (50,301):
        	if(output_img[i][100].all() != 0):
        		count=count+1

        if (count>6):
        	for z in range (1,12):
        		output_img = []
        		output_img = show_webcam(mirror=True)
        		cv2.line(output_img, (100,50), (100,300), (255,0,0), thickness=2, lineType=8, shift=0)
        		cv2.imshow('my webcam', output_img)
        		count1=0
        		for j in range (50,301):
        			if(output_img[j][540].all() != 0):
        				count1=count1+1
        		if(count1>12):
        			cv2.line(output_img,(540,50), (540,300), (255,0,0),thickness=2,lineType=8,shift =0)
        			presentation.SlideShowWindow.View.Next()
        			break
        		else:
        			cv2.line(output_img, (540,50), (540,300) , (0,255,0), thickness=2 , lineType=8 ,shift=0)
        		cv2.imshow('my webcam', output_img)

        else:
        	cv2.line(output_img, (100,50), (100,300), (0,255,0), thickness=2, lineType=8, shift=0)
        	cv2.line(output_img, (540,50), (540,300), (0,255,0), thickness=2, lineType=8, shift=0)

        #Backward movement of hand
        output_img = show_webcam(mirror=True)
        count=0
        for i in range (50,301):
        	if(output_img[i][540].all() != 0):
        		count=count+1
        		print count

        if (count>9):
        	for z in range (1,12):
        		output_img = []
        		output_img = show_webcam(mirror=True)
        		cv2.line(output_img, (540,50), (540,300), (255,0,0), thickness=2, lineType=8, shift=0)
        		cv2.imshow('my webcam', output_img)
        		count1=0
        		for j in range (50,301):
        			if(output_img[j][100].all() != 0):
        				count1=count1+1
        		if(count1>9):
        			cv2.line(output_img,(100,50), (100,300), (255,0,0),thickness=2,lineType=8,shift =0)
        			presentation.SlideShowWindow.View.Previous()
        			break
        		else:
        			cv2.line(output_img, (100,50), (100,300) , (0,255,0), thickness=2 , lineType=8 ,shift=0)
        		cv2.imshow('my webcam', output_img)

        else:
        	cv2.line(output_img, (100,50), (100,300), (0,255,0), thickness=2, lineType=8, shift=0)
        	cv2.line(output_img, (540,50), (540,300), (0,255,0), thickness=2, lineType=8, shift=0)

        cv2.imshow('my webcam', output_img)

        if cv2.waitKey(1) == 27: 
            break  # esc to quit
    cv2.destroyAllWindows()
    app.Quit()

if __name__ == '__main__':
  	main()