# Gesture-Based PowerPoint Controller
This project enables hands-free PowerPoint slide control using computer vision. Built using Python, it uses OpenCV, MediaPipe Hands, and PyWin32 to recognize real-time hand gestures and control slide navigation — no external hardware or clickers needed!


🛠 Technologies Used:  
  * Python  
  * OpenCV – for frame capture and image processing  
  * MediaPipe Hands – for real-time hand tracking and gesture recognition  
  * PyWin32 (win32com) – to control Microsoft PowerPoint via COM automation


🔁 Workflow Overview  -- The system architecture consists of the following stages:
1. Video Capture: OpenCV captures webcam frames.
2. Hand Detection: MediaPipe identifies the hand and its landmarks (21 points).
3. Gesture Recognition: Custom logic checks which fingers are raised.
4. PPT Control: PyWin32 sends commands to PowerPoint based on recognized gestures.



🤚 Gesture Definitions  -- Gesture	Detected Landmark Pattern	Action:  
  * Thumb Up	--->	Go to previous slide
  * Index Finger Up	--->	Go to next slide



                   
# Installation & Setup
Clone this repository:  
  * git clone https://github.com/vidya-reddy01/Gesture-Based-PPT-Controller.git  
  * cd Gesture-Based-PPT-Controller
Install dependencies:  
  * pip install -r requirements.txt  

Run the script:  
  * python Virtual_Mouse.py
    
Ensure your webcam is working and you're in a well-lit environment.




