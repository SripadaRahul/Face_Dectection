import numpy as np
import pandas as pd
import os
import csv
import cv2
from time import sleep
# import face_recognition
from datetime import datetime
import time


def detect(sem,sec):

    if sem == '1' or sem == '2':
        year = 'first_year'
    elif sem == '3' or sem == '4':
        year = 'second_year'
    elif sem == '5' or sem == '6':
        year = 'third_year'
    else:
        year = 'fourth_year'

    filename = f'data/Attendance_xlsx/{year}_{sem}sem_ECE_{sec}.xlsx'

    def from_excel_to_csv():
        df = pd.read_excel(filename, index=False)
        df.to_csv('./data.csv')
    def update_Excel(filename):
        with open('data.csv') as f:
            data = csv.reader(f)
            lines = list(data)

            with open('data.csv', 'w') as g:
                writer = csv.writer(g, lineterminator='\n')
                writer.writerows(lines)

        df = pd.read_csv('data.csv')
        df.to_excel(filename, index=False)
        # print('Attendance is marked in excel')

    path = 'ImagesAttendance'
    images = []
    classNames = []
    myList = os.listdir(path)
    print(myList)
    for cl in myList:
        curImg = cv2.imread(f'{path}/{cl}')
        images.append(curImg)
        classNames.append(os.path.splitext(cl)[0])
    print(classNames)

    def findEncodings(images):
        encodeList = []
        for img in images:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            encode = face_recognition.face_encodings(img)[0]
            encodeList.append(encode)
        return encodeList

    def markAttendance(name):
        with open('data.csv', 'r+') as f:
            myDataList = f.readlines()
            nameList = []
            l=0
            for line in myDataList:
                entry = line.split(',')
                if len(entry[0])>2:
                    l=1
                    nameList.append(entry[0])
            if name not in nameList:
                now = datetime.now()
                ticks = time.localtime(time.time())
                s = list(ticks)
                # tstring=str(s[3]) + ":" + str(s[4]) + ":" + str(s[5])
                dtString = now.strftime('%H:%M:%S')
                if l:
                    f.writelines(f'\n{name},{dtString}')

    #### FOR CAPTURING SCREEN RATHER THAN WEBCAM
    # def captureScreen(bbox=(300,300,690+300,530+300)):
    #     capScr = np.array(ImageGrab.grab(bbox))
    #     capScr = cv2.cvtColor(capScr, cv2.COLOR_RGB2BGR)
    #     return capScr

    encodeListKnown = findEncodings(images)
    print('Encoding Complete')



    cap = cv2.VideoCapture(0)
    lst = []
    while True:
        success, img = cap.read()
        # img = captureScreen()
        imgS = cv2.resize(img, (0, 0), None, 0.25, 0.25)
        imgS = cv2.cvtColor(imgS, cv2.COLOR_BGR2RGB)

        facesCurFrame = face_recognition.face_locations(imgS)
        encodesCurFrame = face_recognition.face_encodings(imgS, facesCurFrame)

        for encodeFace, faceLoc in zip(encodesCurFrame, facesCurFrame):
            matches = face_recognition.compare_faces(encodeListKnown, encodeFace)
            faceDis = face_recognition.face_distance(encodeListKnown, encodeFace)
            # print(faceDis)
            matchIndex = np.argmin(faceDis)

            if matches[matchIndex]:
                name = classNames[matchIndex].upper()
                if name not in lst:
                    lst.append(name)
                    print("Attendance is marked for - "+name)
                y1, x2, y2, x1 = faceLoc
                y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
                cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
                cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
                cv2.putText(img, name, (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)
                markAttendance(name)
                update_Excel(filename)

        cv2.imshow('Webcam', img)
        # if cv2.waitKey(1):
        #     cv2.destroyAllWindows()

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()

# from PIL import ImageGrab

