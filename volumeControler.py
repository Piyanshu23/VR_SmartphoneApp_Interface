import cv2
import mediapipe as mp
import numpy as np
import pyautogui
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
# volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
vol = 0
volBar = 400
volPer = 0

webcam = cv2.VideoCapture(1)
my_hands = mp.solutions.hands.Hands()
drawing_utils = mp.solutions.drawing_utils

while True:
    ret, img = webcam.read()
    img = cv2.flip(img, 1)
    frame_height, frame_width,_ = img.shape
    rgb_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    output = my_hands.process(rgb_img)
    hands = output.multi_hand_landmarks
    if hands:
        for hand in hands:
            drawing_utils.draw_landmarks(img, hand)
            landmarks = hand.landmark
            for id, landmark in enumerate(landmarks):
                x = int(landmark.x * frame_width)
                y = int(landmark.y * frame_height)
                if id == 8:
                    cv2.circle(img, (x,y), 8, (0, 255, 255), 3)
                    x1 = x
                    y1 = y
                if id == 4:
                    cv2.circle(img, (x,y), 8, (0, 0, 255), 3)
                    x2=x
                    y2=y

            # 25-220 range of distance between thumb and first finger
            #-65.5 - 0 : volume range
            cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
            dist = np.hypot(x2-x1,y2-y1)
            #print(dist)
            vol = np.interp(dist, [25, 220], [minVol, maxVol])
            volBar = np.interp(dist, [25, 220], [400, 150])
            volPer = np.interp(dist, [25, 220], [0, 100])
            volume.SetMasterVolumeLevel(vol, None)
            print(int(dist), vol)
    cv2.imshow("Hand detection", img)
    key = cv2.waitKey(10)
    if key == 27:
        break
webcam.release()
cv2.destroyAllWindows()
print(volRange)

