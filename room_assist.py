import pyttsx3 #pip install pyttsx3
import cv2
import datetime
import webbrowser
import os
import numpy as np
from art import tprint
import win32com.client
import speech_recognition as sr

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

#speaker = win32com.client.Dispatch("SAPI.SpVoice") 

def apply_yolo_object_detection(image_to_process, net, out_layers, classes, classes_to_look_for, look_for):
    """
    Recognition and determination of the coordinates of objects on the image
    :param image_to_process: original image
    :return: image with marked objects and captions to them
    """

    height, width, _ = image_to_process.shape
    blob = cv2.dnn.blobFromImage(image_to_process, 1 / 255, (608, 608),
                                 (0, 0, 0), swapRB=True, crop=False)
    net.setInput(blob)
    outs = net.forward(out_layers)
    class_indexes, class_scores, boxes = ([] for i in range(3))
    objects_count = 0

    # Starting a search for objects in an image
    for out in outs:
        for obj in out:
            scores = obj[5:]
            class_index = np.argmax(scores)
            class_score = scores[class_index]
            if class_score > 0:
                center_x = int(obj[0] * width)
                center_y = int(obj[1] * height)
                obj_width = int(obj[2] * width)
                obj_height = int(obj[3] * height)
                box = [center_x - obj_width // 2, center_y - obj_height // 2,
                       obj_width, obj_height]
                boxes.append(box)
                class_indexes.append(class_index)
                class_scores.append(float(class_score))

    # Selection
    chosen_boxes = cv2.dnn.NMSBoxes(boxes, class_scores, 0.0, 0.4)
    for box_index in chosen_boxes:
        box_index = box_index
        box = boxes[box_index]
        class_index = class_indexes[box_index]

        # For debugging, we draw objects included in the desired classes
        if classes[class_index] in classes_to_look_for:
            objects_count += 1
            image_to_process = draw_object_bounding_box(image_to_process,
                                                        class_index, box, classes)

    final_image = draw_object_count(image_to_process, objects_count, look_for)
    return final_image

def draw_object_bounding_box(image_to_process, index, box, classes):
    """
    Drawing object borders with captions
    :param image_to_process: original image
    :param index: index of object class defined with YOLO
    :param box: coordinates of the area around the object
    :return: image with marked objects
    """

    x, y, w, h = box
    start = (x, y)
    end = (x + w, y + h)
    color = (0, 255, 0)
    width = 2
    final_image = cv2.rectangle(image_to_process, start, end, color, width)

    start = (x, y - 10)
    font_size = 1
    font = cv2.FONT_HERSHEY_SIMPLEX
    width = 2
    text = classes[index]
    final_image = cv2.putText(final_image, text, start, font,
                              font_size, color, width, cv2.LINE_AA)

    return final_image

def draw_object_count(image_to_process, objects_count, look_for):
    """
    Signature of the number of found objects in the image
    :param image_to_process: original image
    :param objects_count: the number of objects of the desired class
    :return: image with labeled number of found objects
    """

    start = (10, 120)
    font_size = 1.5
    font = cv2.FONT_HERSHEY_SIMPLEX
    width = 3
    text = "Objects found: " + str(objects_count)
    #s=("selected object found")
    #speak(s)
    #s=(objects_count)
    #speak(s)
    #print(objects_count)
    # Text output with a stroke
    # (so that it can be seen in different lighting conditions of the picture)
    white_color = (255, 255, 255)
    black_outline_color = (0, 0, 0)
    final_image = cv2.putText(image_to_process, text, start, font, font_size,
                              black_outline_color, width * 3, cv2.LINE_AA)
    final_image = cv2.putText(final_image, text, start, font, font_size,
                              white_color, width, cv2.LINE_AA)
    if(objects_count>0):
      s=("selected object found")
      speak(s)
      s=(look_for)
      speak(s)

    return final_image   
    
def start_video_object_detection(net, layer_names, out_layers_indexes, out_layers, classes, classes_to_look_for, look_for):
    while True:
        try:
            # Capturing a picture from a video
            video_camera_capture = cv2.VideoCapture(0)
            
            while video_camera_capture.isOpened():
                ret, frame = video_camera_capture.read()
                if not ret:
                    break
                
                # Application of object recognition methods on a video frame from YOLO
                frame = apply_yolo_object_detection(frame, net, 
                    out_layers, classes, classes_to_look_for, look_for)
                
                # Displaying the processed image on the screen with a reduced window size
                frame = cv2.resize(frame, (1920 // 2, 1080 // 2))
                cv2.imshow("Video Capture", frame)
                
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
            break
    
        except KeyboardInterrupt:
            print("Keyboard Interrupt detected. Exiting...")
            exit()

        finally:
            # Release resources properly
            if video_camera_capture is not None:
                video_camera_capture.release()
            cv2.destroyAllWindows()
    return

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Friday. How may I help you?")       

def takeCommand():
    #It takes microphone input from the user and returns string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        # print(e)    
        speak("Say that again please...")  
        return "None"
    return query

def find_things(MyText):    
    # Logo
    tprint("Object detection")

    # Loading YOLO scales from files and setting up the network
    net = cv2.dnn.readNetFromDarknet("Resources/yolov4-tiny.cfg",
                                     "Resources/yolov4-tiny.weights")
    layer_names = net.getLayerNames()
    out_layers_indexes = net.getUnconnectedOutLayers()
    out_layers = [layer_names[index - 1] for index in out_layers_indexes]

    # Loading from a file of object classes that YOLO can detect
    with open("Resources/coco.names.txt") as file:
        classes = file.read().split("\n")

    # Determining classes that will be prioritized for search in an image
    # The names are in the file coco.names.txt

    #MyText = takeCommand()
    
    #look_for = input("What we are looking for: ").split(',')
    look_for = (MyText).split(',')
    # Delete spaces
    list_look_for = []
    for look in look_for:
        list_look_for.append(look.strip())

        classes_to_look_for = list_look_for

        start_video_object_detection(net, layer_names, out_layers_indexes,
                                    out_layers, classes, classes_to_look_for, look_for)

if __name__ == '__main__':
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()
        # Logic for executing tasks based on query
        #Starts YouTube
        if 'YouTube' in query:
            webbrowser.open("youtube.com")
        # Plays music
        elif 'music' in query:
            music_dir = 'C:/Users/HP/Music'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[0]))
        # Tells time
        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
        # Finds objects
        elif 'find' in query:
            speak("What should I find?")
            thing = takeCommand().lower()
            find_things(thing)
        # Opens camera
        elif 'open camera' in query:
            cap = cv2.VideoCapture(0)
            if not cap.isOpened():
                print("Error: Could not open camera.")
                continue
            while True:
                ret, frame = cap.read()
                cv2.imshow('Camera', frame)
                q = takeCommand().lower()
                if 'close camera' in q:
                    break
                elif 'how do i look' in q:
                    speak("You are looking very attractive today, sir!")
            cap.release()
            cv2.destroyAllWindows()
        # Greetings1
        elif 'how are you' in query:
            speak("I'm Fine. How are you? Hope you are enjoying your day!")
        # Greetings2
        elif 'nice' in query:
            speak("Okay, let me know if you need any help.")
        # Introduces itself
        elif 'who are you' in query:
            speak("I'm Female Replacement Intelligence Digital Assistant for Youth.")
        # Stops working and exits
        elif 'exit' in query:
            speak("exiting now! Have a nice day! Bye!")
            exit()
