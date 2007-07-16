from System.IO import FileSystemWatcher, Path
from System.Threading import Thread

from UdpSender import UdpSender, port, group

directory = Path.GetDirectoryName(Path.GetFullPath(__file__))
dataFile =  'data.txt'

watcher = FileSystemWatcher()
watcher.Path = directory
watcher.Filter = dataFile

sender = UdpSender(port, group)

def onChanged(source, event):
    print 'Changed:', event.ChangeType, event.FullPath
    sender.send('DATA:FILE:CHANGED')

watcher.Changed += onChanged

watcher.EnableRaisingEvents = True 

while True:
    Thread.Sleep(1000)
    