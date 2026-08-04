# services/camera_service.py
import cv2
import numpy as np
from pyzbar.pyzbar import decode
from services.qr_service import QRService

class CameraService:
    @staticmethod
    def get_camera(camera_index=0):
        cap = cv2.VideoCapture(camera_index)
        if not cap.isOpened():
            return None
        # Set resolution
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        return cap

    @staticmethod
    def process_frame(frame):
        """Resizes frame to (640, 480) and attempts to find/decode QR code.
        Returns: (processed_frame, decoded_data_str or None)
        """
        # Resize frame
        resized = cv2.resize(frame, (640, 480))
        gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
        
        # Threshold
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        decoded_objects = decode(resized)
        if not decoded_objects:
            decoded_objects = decode(thresh)
            
        decoded_str = None
        for obj in decoded_objects:
            # Draw boundary polygon
            points = obj.polygon
            if len(points) > 4:
                hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                cv2.polylines(resized, [hull], True, (0, 255, 0), 2)
            elif len(points) > 0:
                cv2.polylines(resized, [np.array(points, dtype=np.int32)], True, (0, 255, 0), 2)
            
            try:
                raw_data = obj.data.decode('utf-8')
                decoded_str = QRService.decode_qr(raw_data)
                break # Only process one QR code per frame
            except Exception:
                pass
                
        return resized, decoded_str
