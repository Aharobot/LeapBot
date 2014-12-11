# imports
import Leap, sys, time, serial, win32com.client
from Leap import CircleGesture, KeyTapGesture, ScreenTapGesture, SwipeGesture

# Serial port parameters
serial_speed = 9600
# Your serial port instead of ttyUSB0
serial_port = 'COM6'
ser = serial.Serial(serial_port, serial_speed, timeout=1)

# Initializing voice output
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Speak("I'm Robo-51.I'm at your command!!")

class SampleListener(Leap.Listener):
    alarm = 0
    def on_init(self, controller):
        pass
        #print ""

    def on_connect(self, controller):
        print "Connected"

        # Enable gestures
        controller.enable_gesture(Leap.Gesture.TYPE_CIRCLE);
        controller.enable_gesture(Leap.Gesture.TYPE_KEY_TAP);
        controller.enable_gesture(Leap.Gesture.TYPE_SCREEN_TAP);
        controller.enable_gesture(Leap.Gesture.TYPE_SWIPE);

    def on_disconnect(self, controller):
        # Note: not dispatched when running in a debugger.
        print "Disconnected"

    def on_exit(self, controller):
        print "Exited"

    def on_frame(self, controller):
        # Get the most recent frame and report some basic information
        frame = controller.frame()
        if not frame.hands.is_empty: 
            for gesture in frame.gestures():
                if gesture.type == Leap.Gesture.TYPE_SWIPE:
                    swipe = SwipeGesture(gesture)

                    if swipe.direction[0] > 0.5 and self.alarm < time.time():
                        self.alarm = time.time() + 1.00
                        print "Right"
                        #voice output
                        speaker.Speak("Right!")
                        #write data to robot
                        ser.write('b')
                        
                    if swipe.direction[0] < -0.5 and self.alarm < time.time():
                        self.alarm = time.time() + 1.00
                        print "Left"
                        #voice output
                        speaker.Speak("Left!)
                        #write data to robot
                        ser.write('c')
                        
                    if swipe.direction[1] > 0.5 and self.alarm < time.time():
                        self.alarm = time.time() + 1.00
                        print "Forward"
                        #voice output
                        speaker.Speak("Forward!")
                        #write data to robot
                        ser.write('d')
                        
                    if swipe.direction[1] < -0.5 and self.alarm < time.time():
                        self.alarm = time.time() + 1.00
                        print "Backward"
                        #voice output
                        speaker.Speak("Backward!")
                        #write data to robot
                        ser.write('a')
                        
                if gesture.type == Leap.Gesture.TYPE_SCREEN_TAP and self.alarm < time.time():
                        self.alarm = time.time() + 1.00
                        print "Stop"
                        #voice output
                        speaker.Speak("Stop!")
                        #write data to robot
                        ser.write('s')
                        
                if gesture.type == Leap.Gesture.TYPE_KEY_TAP and self.alarm < time.time():
                        self.alarm = time.time() + 1.00
                        print "Stop"
                        #voice output
                        speaker.Speak("Stop!")
                        #write data to robot
                        ser.write('s')

    def state_string(self, state):
        if state == Leap.Gesture.STATE_START:
            return "STATE_START"

        if state == Leap.Gesture.STATE_UPDATE:
            return "STATE_UPDATE"

        if state == Leap.Gesture.STATE_STOP:
            return "STATE_STOP"

        if state == Leap.Gesture.STATE_INVALID:
            return "STATE_INVALID"

def main():
    print "swipe right for turning right"
    print "swipe left for turning left"
    print "swipe up for going forward"
    print "swipe down for going backward"
    print "key/screen tap for stopping"
    print ""

    # Create a sample listener and controller
    listener = SampleListener()
    controller = Leap.Controller()

    # Have the sample listener receive events from the controller
    controller.add_listener(listener)

    # Keep this process running until Enter is pressed
    #print "Press Enter to quit..."
    sys.stdin.readline()

    # Remove the sample listener when done
    controller.remove_listener(listener)


if __name__ == "__main__":
    main()