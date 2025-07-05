import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import aspose.slides as slides
Application = win32com.client.Dispatch("PowerPoint.Application" )
Presentation = Application.Presentations.Open("D:\Downloads\Mini project 025&026.pptx")
print(Presentation.Name)
Presentation.SlideShowSettings.Run()
# Parameters
width, height = 900, 720
gestureThreshold = 500
# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height) 
# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)
# Variables
imgList = []
delay = 30
buttonPressed = False
counter = 0
imgNumber = 20
while True:
    # Get image frame
    success, img = cap.read()
    # Find the hand and its landmarks
    hands, img = detectorHand.findHands(img)  # with draw
    if hands and buttonPressed is False:  # If hand is detected
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]  # List of 21 Landmark points
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
        if cy <= gestureThreshold:  # If hand is at the height of the face
            if fingers == [0, 1, 0, 0, 0]:
                print(imgNumber)
                print("Next")
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Next()
                    imgNumber += 1
            if fingers == [1, 0, 0, 0, 0]:
                print(imgNumber)
                print("Previous")
                buttonPressed = True
                if imgNumber >0 :
                    Presentation.SlideShowWindow.View.Previous()
                    imgNumber -= 1

 
    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False
 

 
    cv2.imshow("Image", img)
 
    key = cv2.waitKey(1)
    if key == ord('q'):
        break