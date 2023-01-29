
import psutil

for proc in psutil.process_iter():
    # name = proc.name()
    # print(name)
    if "main" in proc.name():
        if proc.children():
            for sub in proc.children():
                print(sub.kill())
        print(proc.kill())
                # if "chromedriver" in sub.name():
                #     print(f'{proc.name()}: pid {proc.pid}: {proc.kill()}')
