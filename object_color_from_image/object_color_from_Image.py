import os
import time
import cv2
import numpy as np
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from skimage import io, color
from sklearn.cluster import KMeans
from collections import Counter

BASE_PATH = "G:\\Git_2\\python_tools\\object_color_from_image"

# Pre-trained MobileNet SSD model for object detection
net = cv2.dnn.readNetFromCaffe(os.path.join(BASE_PATH, 'MobileNetSSD_deploy.prototxt'), os.path.join(BASE_PATH, 'MobileNetSSD_deploy.caffemodel'))
CLASSES = ["background", "aeroplane", "bicycle", "bird", "boat", "bottle", "bus", "car", "cat", "chair", "cow", "diningtable", "dog", "horse", "motorbike", "person", "pottedplant", "sheep", "sofa", "train", "tvmonitor"]

def detect_objects_and_colors(image_path):
    image = cv2.imread(image_path)
    (h, w) = image.shape[:2]
    
    # Object Detection
    blob = cv2.dnn.blobFromImage(image, 0.007843, (300, 300), 127.5)
    net.setInput(blob)
    detections = net.forward()

    detected_objects = []
    for i in range(detections.shape[2]):
        confidence = detections[0, 0, i, 2]
        if confidence > 0.6:
            idx = int(detections[0, 0, i, 1])
            detected_objects.append(CLASSES[idx])

    # Color Detection (top 5 colors)
    image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    image_rgb = image_rgb.reshape((image_rgb.shape[0] * image_rgb.shape[1], 3))
    clt = KMeans(n_clusters=5, n_init=10)
    labels = clt.fit_predict(image_rgb)
    label_counts = Counter(labels)
    dominant_colors = clt.cluster_centers_

    # Save metadata
    with open(os.path.join(BASE_PATH, 'output', os.path.basename(image_path) + ".meta"), 'w') as f:
        f.write("Detected Objects:\n")
        for obj in detected_objects:
            f.write(f"{obj}\n")

        f.write("\nTop 5 Dominant Colors (RGB):\n")
        for color in dominant_colors:
            f.write(f"{int(color[0])}, {int(color[1])}, {int(color[2])}\n")

class ImageHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.is_directory:
            return None

        elif event.event_type == 'modified':
            # Check if image
            if event.src_path.lower().endswith(('.png', '.jpeg', '.jpg', '.tiff')):
                detect_objects_and_colors(event.src_path)

if __name__ == "__main__":
    path = os.path.join(BASE_PATH, 'input')
    event_handler = ImageHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
