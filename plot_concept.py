import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import cv2

def main():
    crop_image("./test.png")

def crop_image(image):
    img = cv2.imread(image) # Read in the image and convert to grayscale
    #img = img[:-20,:-20] # Perform pre-cropping
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = 255*(gray < 128).astype(np.uint8) # To invert the text to white
    gray = cv2.morphologyEx(gray, cv2.MORPH_OPEN, np.ones((2, 2), dtype=np.uint8)) # Perform noise filtering
    coords = cv2.findNonZero(gray) # Find all non-zero points (text)
    x, y, w, h = cv2.boundingRect(coords) # Find minimum spanning bounding box
    x = 0
    y = 0
    h = h + 20
    rect = img[y:y+h, x:x+w] # Crop the image - note we do this on the original image
    cv2.imshow("Cropped", rect) # Show it
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    cv2.imwrite("rect.png", rect) # Save the image


if __name__ == '__main__':
    main()