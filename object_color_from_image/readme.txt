The object detection uses a MobileNet SSD model. You'll need to download the MobileNetSSD_deploy.prototxt and MobileNetSSD_deploy.caffemodel files. You can find them in the official GitHub repository of Caffe.
This script uses the cv2 (OpenCV) and watchdog libraries, which need to be installed via pip (pip install opencv-python-headless watchdog).
Object detection threshold is set to 0.2 (20% confidence). Adjust as needed.
For color detection, this script uses the KMeans clustering algorithm. You'll need to install skimage for that (pip install scikit-image).
Place this script in the ...\\object_color_from_image directory and run it. It will monitor the input directory for new images. Once you add an image, it will process the image and create a metadata file in the output directory.