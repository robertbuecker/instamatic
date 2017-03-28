from Tkinter import *
from PIL import Image, ImageEnhance
from PIL import ImageTk
import threading
import numpy as np
import time
import datetime
from instamatic.formats import write_tiff
from camera import Camera


class VideoStreamer(object):
    """docstring for VideoStreamer"""
    def __init__(self, cam, handler):
        super(VideoStreamer, self).__init__()
        
        self.handler = handler
        self.cam = cam

        self.default_exposure = self.cam.default_exposure
        self.default_binsize = self.cam.default_binsize
        self.dimensions = self.cam.dimensions
        self.defaults = self.cam.defaults
        self.name = self.cam.name

        self.frame = None
        self.thread = None
        self.stopEvent = None

        self.stash = None

        self.frametime = 0.1
        self.exposure = self.frametime
        self.binsize = self.cam.default_binsize

        self.lock = threading.Lock()

        self.stopEvent = threading.Event()
        self.acquireInitiateEvent = threading.Event()
        self.acquireCompleteEvent = threading.Event()

    def getImage(self, t=None, binsize=1):
        if not t:
            t = self.exposure

        self.lock.acquire(True)
        img = np.ones((512, 512))*256
        # self.cam.getImage(t=t, binsize=1, fastmode=True)
        time.sleep(t)
        self.lock.release()
        # self.handler.send_frame(img)
        return img

    def run(self):
        while not self.stopEvent.is_set():
            if self.acquireInitiateEvent.is_set():
                self.acquireInitiateEvent.clear()
                frame = self.getImage()
                self.handler.send_acquire(frame)
            else:
                frame = np.random.random((512,512)) * 256
                time.sleep(self.frametime + 0.01)
                # frame = self.cam.getImage(t=self.frametime, fastmode=True)
                self.handler.send_frame(frame)

    def start_loop(self):
        self.thread = threading.Thread(target=self.run, args=())
        self.thread.start()

    def end_loop(self):
        self.thread.stop()


class VideoViewer(threading.Thread):
    """docstring for VideoViewer"""
    def __init__(self, cam="simulate"):
        threading.Thread.__init__(self)

        self.cam = Camera(kind=cam)
        self.stream = self.setup_stream()

        self.panel = None

        self.frametime = 0.05
        self.contrast = 1.0

        self.last = time.time()
        self.nframes = 1
        self.update_frequency = 0.5

        self.start()

    def run(self):
        self.root = Tk()

        self.init_vars()
        self.buttonbox(self.root)
        self.header(self.root)
        self.makepanel(self.root)

        # self.stopEvent = threading.Event()
 
        self.root.wm_title("Instamatic stream")
        self.root.wm_protocol("WM_DELETE_WINDOW", self.close)

        self.root.bind('<Escape>', self.close)

        self.root.bind('<<StreamAcquire>>', self.on_frame)
        self.root.bind('<<StreamEnd>>', self.close)
        self.root.bind('<<StreamFrame>>', self.on_frame)

        self.start_stream()
        self.root.mainloop()

    def header(self, master):
        ewidth = 10
        lwidth = 12

        frame = Frame(master)
        self.e_fps       = Entry(frame, bd=0, width=ewidth, textvariable=self.var_fps, state=DISABLED)
        self.e_frametime = Entry(frame, bd=0, width=ewidth, textvariable=self.var_frametime, state=DISABLED)
        self.e_overhead = Entry(frame, bd=0, width=ewidth, textvariable=self.var_overhead, state=DISABLED)
        
        Label(frame, anchor=E, width=lwidth, text="fps:").grid(row=1, column=0)
        self.e_fps.grid(row=1, column=1)
        Label(frame, anchor=E, width=lwidth, text="frametime (ms):").grid(row=1, column=2)
        self.e_frametime.grid(row=1, column=3)
        Label(frame, anchor=E, width=lwidth, text="overhead (ms):").grid(row=1, column=4)
        self.e_overhead.grid(row=1, column=5)
        
        frame.pack()

        frame = Frame(master)
        
        self.e_exposure = Spinbox(frame, width=ewidth, textvariable=self.var_exposure, from_=0.0, to=1.0, increment=0.01)
        
        Label(frame, anchor=E, width=lwidth, text="exposure (s)").grid(row=1, column=0)
        self.e_exposure.grid(row=1, column=1)

        self.e_contrast = Spinbox(frame, width=ewidth, textvariable=self.var_contrast, from_=0.0, to=10.0, increment=0.1)
        
        Label(frame, anchor=E, width=lwidth, text="Contrast").grid(row=1, column=2)
        self.e_contrast.grid(row=1, column=3)
        
        frame.pack()

    def makepanel(self, master, resolution=(512,512)):
        if self.panel is None:
            image = Image.fromarray(np.zeros(resolution))
            image = ImageTk.PhotoImage(image)

            self.panel = Label(image=image)
            self.panel.image = image
            self.panel.pack(side="left", padx=10, pady=10)

    def buttonbox(self, master):
        btn = Button(master, text="Save image",
            command=self.saveImage)
        btn.pack(side="bottom", fill="both", expand="yes", padx=10, pady=10)

    def init_vars(self):
        self.var_fps = DoubleVar()
        self.var_frametime = DoubleVar()
        self.var_overhead = DoubleVar()

        self.var_exposure = DoubleVar()
        self.var_exposure.set(self.cam.default_exposure)
        self.var_exposure.trace("w", self.update_exposure_time)

        self.var_contrast = DoubleVar(value=1.0)
        self.var_contrast.set(self.contrast)
        self.var_contrast.trace("w", self.update_contrast)

    def update_exposure_time(self, name, index, mode):
        # print name, index, mode
        try:
            self.frametime = self.var_exposure.get()
        except:
            pass
        else:
            self.stream.frametime = self.frametime

    def update_contrast(self, name, index, mode):
        # print name, index, mode
        try:
            self.contrast = self.var_contrast.get()
        except:
            pass

    def saveImage(self):
        outfile = datetime.datetime.now().strftime("%Y%m%d-%H%M%S.%f") + ".tiff"
        write_tiff(outfile, self.frame)
        print " >> Wrote file:", outfile

    def close(self):
        self.stream.stopEvent.set()
        self.root.quit()

    def setup_stream(self):
        class Handler(object):
            def need_stop(self_):
                return self.stream.stopEvent.is_set()

            def send_frame(self_, frame):
                self.stream.lock.acquire(True)
                self.frame = frame
                self.stream.lock.release()

                self.root.event_generate('<<StreamFrame>>', when='tail')

            def send_acquire(self_, acquiredFrame):
                self.stream.lock.acquire(True)
                self.acquiredFrame = self.frame = acquiredFrame
                self.stream.lock.release()
                self.stream.acquireCompleteEvent.set()

                self.root.event_generate('<<StreamFrame>>', when='tail')

        # self.cam.stopEvent.clear()
        return VideoStreamer(self.cam, Handler())
    
    def start_stream(self):
        self.stream.start_loop()

    def on_frame(self, event):
        self.stream.lock.acquire(True)
        frame = self.frame
        self.stream.lock.release()

        if self.contrast != 1:
            image = Image.fromarray(frame).convert("L")
            image = ImageEnhance.Contrast(image).enhance(self.contrast)
            # Can also use ImageEnhance.Sharpness or ImageEnhance.Brightness if needed
        else:
            image = Image.fromarray(frame)

        image = ImageTk.PhotoImage(image=image)

        self.panel.configure(image=image)
        # keep a reference to avoid premature garbage collection
        self.panel.image = image

        self.update_frametimes()
        self.root.update_idletasks()

    def update_frametimes(self):
        self.current = time.time()
        delta = self.current - self.last

        if delta > self.update_frequency:
            frametime = delta/self.nframes
            fps = 1.0/frametime
            overhead = frametime - self.stream.frametime

            self.var_fps.set(round(fps, 2))
            self.var_frametime.set(round(frametime*1000, 2))
            self.var_overhead.set(round(overhead*1000, 2))
            self.last = self.current
            self.nframes = 1
        else:
            self.nframes += 1

    def getImage(self, exposure=None, binsize=1):
        current_frametime = self.stream.frametime

        # set to 0 to prevent it lagging the acquisition
        self.stream.frametime = 0
        self.stream.exposure = exposure
        self.stream.binsize = binsize

        self.stream.acquireInitiateEvent.set()
        self.stream.acquireCompleteEvent.wait()

        self.stream.lock.acquire(True)
        frame = self.acquiredFrame
        self.stream.lock.release()
        
        self.stream.acquireCompleteEvent.clear()
        self.stream.frametime = current_frametime
        
        return frame


if __name__ == '__main__':
    # stream = VideoStream(cam="simulate")
    stream = VideoViewer(cam="simulate")
    # stream.root.mainloop()
    from IPython import embed
    embed()
    stream.close()