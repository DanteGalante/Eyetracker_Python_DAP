#  python3 -m venv work
# source work/bin/activate
# pip install opencv-python
import cv2
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import get_column_letter



cap = cv2.VideoCapture("eye_recording.flv")


wb = Workbook()
dest_filename = 'coordenadas.xlsx'
ws1 = wb.active
ws1.title = "Coordenadas"
headers = ['Eje X', 'Eje Y', 'Tiempo en segundos', 'Tiempo en milisegundos']
ws1.append(headers)

fps = cap.get(cv2.CAP_PROP_FPS)


while True:
    ret, frame = cap.read()
    if ret is False:
        break

    timestamps = [cap.get(cv2.CAP_PROP_POS_MSEC)]
    timeSeconds = (round((timestamps[0] / 1000), 2))

    roi = frame[369: 805, 900: 1700]
    rows, cols, _ = roi.shape
    gray_roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    gray_roi = cv2.GaussianBlur(gray_roi, (7, 7), 0)

    _, threshold = cv2.threshold(gray_roi, 3, 255, cv2.THRESH_BINARY_INV)
    contours, hierarchy = cv2.findContours(threshold, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    contours = sorted(contours, key=lambda x: cv2.contourArea(x), reverse=True)

    for cnt in contours:
        (x, y, w, h) = cv2.boundingRect(cnt)

        #cv2.drawContours(roi, [cnt], -1, (0, 0, 255), 3)
        cv2.rectangle(roi, (x, y), (x + w, y + h), (255, 0, 0), 2)
        cv2.line(roi, (x + int(w/2), 0), (x + int(w/2), rows), (0, 255, 0), 2)
        cv2.line(roi, (0, y + int(h/2)), (cols, y + int(h/2)), (0, 255, 0), 2)
        center = (x + w // 2, y + h // 2)
        radius = 2
        cv2.circle(roi, center, radius, (255, 255, 0), 2)


        rowData = [x + w // 2, y + h // 2, timeSeconds, timestamps[0]]
        ws1.append(rowData)
        print(rowData)

        break


    cv2.imshow("Threshold", threshold)
    cv2.imshow("gray roi", gray_roi)
    cv2.imshow("Roi", roi)
    key = cv2.waitKey(30)
    if key == 27:
        break

wb.save(filename = dest_filename)

cv2.destroyAllWindows()
