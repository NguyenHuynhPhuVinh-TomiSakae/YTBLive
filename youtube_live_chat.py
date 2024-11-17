from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import os
import time
import google.generativeai as genai
from datetime import datetime
import win32gui
import win32con
import win32api
import win32clipboard
from dotenv import load_dotenv

class YouTubeChatBot:
    def __init__(self):
        self.youtube = self.get_youtube_service()
        self.live_chat_id = None
        # Khởi tạo Gemini
        load_dotenv()  # Load biến môi trường từ file .env
        gemini_api_key = os.getenv('GEMINI_API_KEY')
        if not gemini_api_key:
            raise ValueError("Không tìm thấy GEMINI_API_KEY trong file .env")
        genai.configure(api_key=gemini_api_key)
        self.model = genai.GenerativeModel('gemini-1.5-flash')
        self.notepad_hwnd = None
        self.find_notepad()

    def get_youtube_service(self):
        """Xác thực và tạo YouTube service"""
        SCOPES = ['https://www.googleapis.com/auth/youtube.force-ssl']
        creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('live.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        return build('youtube', 'v3', credentials=creds)

    def get_live_chat_id(self):
        """Lấy live chat ID từ stream đang hoạt động"""
        try:
            broadcasts = self.youtube.liveBroadcasts().list(
                part="snippet",
                broadcastStatus="active",
                maxResults=1
            ).execute()

            if broadcasts.get('items'):
                return broadcasts['items'][0]['snippet']['liveChatId']
            return None
        except Exception as e:
            print(f"Lỗi khi lấy live chat ID: {e}")
            return None

    def find_notepad(self):
        """Tìm cửa sổ Notepad đang mở"""
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if 'Notepad' in window_text:
                    hwnds.append(hwnd)
            return True
        
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        if hwnds:
            self.notepad_hwnd = hwnds[0]
            print("Đã tìm thấy cửa sổ Notepad")
        else:
            print("Không tìm thấy cửa sổ Notepad đang mở!")

    def write_to_notepad(self, text):
        """Ghi text vào Notepad đang mở"""
        if not self.notepad_hwnd:
            print("Không tìm thấy cửa sổ Notepad!")
            self.find_notepad()
            if not self.notepad_hwnd:
                return
            
        try:
            # Đưa text vào clipboard với encoding UTF-16
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            text_utf16 = text.encode('utf-16-le')
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text_utf16)
            win32clipboard.CloseClipboard()
            
            # Thử focus vào Notepad nhiều lần
            max_attempts = 3
            for _ in range(max_attempts):
                try:
                    if not win32gui.IsWindow(self.notepad_hwnd):
                        print("Cửa sổ Notepad đã đóng, đang tìm lại...")
                        self.find_notepad()
                        if not self.notepad_hwnd:
                            return
                    
                    win32gui.ShowWindow(self.notepad_hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(self.notepad_hwnd)
                    time.sleep(0.1)
                    
                    # Paste và xuống dòng
                    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                    win32api.keybd_event(ord('V'), 0, 0, 0)
                    win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
                    time.sleep(0.1)
                    win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                    win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                    break
                except Exception as e:
                    print(f"Lần thử {_ + 1}: Không thể focus vào Notepad, đang thử lại...")
                    time.sleep(0.5)
            
        except Exception as e:
            print(f"Lỗi khi ghi vào Notepad: {e}")

    def process_message(self, message):
        """Xử lý tin nhắn với AI Gemini và ghi vào Notepad"""
        author = message['authorDetails']['displayName']
        text = message['snippet']['displayMessage']
        print(f"Tin nhắn từ {author}: {text}")
        
        # Xử lý với Gemini
        response = self.model.generate_content(text)
        ai_response = response.text
        
        # Chuẩn bị nội dung để ghi
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        content = f"\n[{timestamp}]\nUser ({author}): {text}\nAI: {ai_response}\n{'-' * 50}\n"
        
        # Ghi vào Notepad
        self.write_to_notepad(content)

    def run(self):
        """Chạy bot"""
        print("Đang khởi động bot...")
        self.live_chat_id = self.get_live_chat_id()
        
        if not self.live_chat_id:
            print("Không tìm thấy live chat đang hoạt động!")
            return

        print(f"Đã kết nối với live chat ID: {self.live_chat_id}")
        next_page_token = None

        while True:
            try:
                chat_messages = self.youtube.liveChatMessages().list(
                    liveChatId=self.live_chat_id,
                    part="snippet,authorDetails",
                    pageToken=next_page_token
                ).execute()

                next_page_token = chat_messages.get('nextPageToken')

                for message in chat_messages['items']:
                    # Chỉ xử lý tin nhắn với AI và ghi vào Notepad
                    self.process_message(message)

                time.sleep(5)  # Đợi 5 giây trước khi kiểm tra tin nhắn mới

            except Exception as e:
                print(f"Lỗi: {e}")
                time.sleep(5)

def main():
    bot = YouTubeChatBot()
    bot.run()

if __name__ == "__main__":
    main()