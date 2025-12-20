from kivy.lang import Builder
from kivymd.app import MDApp
from kivymd.uix.pickers import MDTimePicker
from datetime import datetime
from kivy.core.window import Window

Window.size = (700, 500)

KV = """ 
MDFloatLayout:

    MDRaisedButton:
        text: "Open Time Picker"
        pos_hint: {'center_x': .5, "center_y": .5}
        on_release: app.show_time_picker()

"""

class TimePickerApp(MDApp):
    def build(self):
        return Builder.load_string(KV)
    
    #Get Time
    def get_time(self, instance, time):
        print(time.strftime("%H:%M"), flush = True)
        self.stop()

    # Cancel 
    def on_cancel(self, instance, time):
        print("", flush = True)
        self.stop()
    
    def show_time_picker(self):
        default_time = datetime.now()
        dialog = MDTimePicker()

        #Set default time
        dialog.set_time(default_time)

        dialog.bind(on_save = self.get_time, on_cancel = self.on_cancel)
        dialog.open()

if __name__ == "__main__":
    TimePickerApp().run()
