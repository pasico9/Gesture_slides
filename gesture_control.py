import cv2
import mediapipe as mp
import pyautogui
import time

cap = cv2.VideoCapture(0)

mp_hands = mp.solutions.hands
hands = mp_hands.Hands(
    min_detection_confidence=0.7,
    min_tracking_confidence=0.7,
    max_num_hands=1
)

mp_draw = mp.solutions.drawing_utils

last_action = 0
COOLDOWN = 1.2

prev_gesture = None

while True:

    ret, frame = cap.read()
    frame = cv2.flip(frame, 1)

    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    result = hands.process(rgb)

    gesture = None

    if result.multi_hand_landmarks:

        for hand_landmarks in result.multi_hand_landmarks:

            mp_draw.draw_landmarks(frame, hand_landmarks,
                                   mp_hands.HAND_CONNECTIONS)

            lm = hand_landmarks.landmark

            thumb_tip = lm[4]
            thumb_ip = lm[3]

            index_tip = lm[8]
            index_pip = lm[6]

            middle_tip = lm[12]
            middle_pip = lm[10]

            ring_tip = lm[16]
            ring_pip = lm[14]

            pinky_tip = lm[20]
            pinky_pip = lm[18]

            thumb_up = thumb_tip.y < thumb_ip.y
            index_up = index_tip.y < index_pip.y
            middle_up = middle_tip.y < middle_pip.y
            ring_up = ring_tip.y < ring_pip.y
            pinky_up = pinky_tip.y < pinky_pip.y

            if thumb_up and not index_up and not middle_up and not ring_up and not pinky_up:
                gesture = "thumb"

            elif index_up and not middle_up and not ring_up and not pinky_up and not thumb_up:
                gesture = "index"

    current_time = time.time()

    if gesture != prev_gesture and (current_time - last_action > COOLDOWN):

        if gesture == "thumb":
            pyautogui.press("right")
            print("Next Slide")
            last_action = current_time

        elif gesture == "index":
            pyautogui.press("left")
            print("Previous Slide")
            last_action = current_time

    prev_gesture = gesture

    cv2.imshow("Gesture Control", frame)

    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

cap.release()
cv2.destroyAllWindows()