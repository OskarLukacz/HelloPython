import sched
import time
import atexit
import sys
import xlsxwriter
import pytz
from datetime import datetime
from pubnub.callbacks import SubscribeCallback
from pubnub.enums import PNStatusCategory
from pubnub.pnconfiguration import PNConfiguration
from pubnub.pubnub import PubNub

cycle = 0
count = 5
interval = 5
shut_down = False

allEvents = []
data = {
    "ping": ['empty'],
    "ack": ['empty'],
    "upTime": ['empty'],
    "rawPing": ['empty'],
    "rawAck": ['empty']
}

pnconfig = PNConfiguration()
pnconfig.subscribe_key = 'sub-c-05dce56c-3c2e-11e7-847e-02ee2ddab7fe'
pnconfig.publish_key = 'pub-c-b6db3020-95a8-4c60-8d16-13345aaf8709'

pubnub = PubNub(pnconfig)
s = sched.scheduler(time.time, time.sleep)

pubnubEvents = []

def my_publish_callback(envelope, status):
    # Check whether request successfully completed or not
    if not status.is_error():
        pass  # Message successfully published to specified channel.
    else:
        pass  # Handle message publish error. Check 'category' property to find out possible issue
        # because of which request did fail.
        # Request can be resent using: [status retry];


class MySubscribeCallback(SubscribeCallback):

    def status(self, pubnub, status):
        if status.category == PNStatusCategory.PNConnectedCategory:
            pubnub.publish().channel("hello_world").message("Python script hello.py on channel hello_world...").pn_async(my_publish_callback)

    def message(self, pubnub, message):
        pass
        t = datetime.now(pytz.timezone('US/Central')).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        print(message.message)
        handle_message(str(message.message), t)


def handle_message(raw_message, timestamp):
    stamped_message = "at: " + timestamp + "    " + raw_message
    allEvents.append(stamped_message)
    if str(raw_message[2]) == 'x':
        global shut_down
        shut_down = True
    if raw_message[2] == 'p':
        if raw_message[10:18] == "hello.py":
            data["ping"][cycle] = timestamp
            data["rawPing"][cycle] = stamped_message
    if raw_message[2] == 's':
        if raw_message[12:19] == "pingAck":
            cur_uptime = raw_message[raw_message.find('T')+1:raw_message.find('}')-1]
            data["upTime"][cycle] = cur_uptime
            data["ack"][cycle] = timestamp
            data["rawAck"][cycle] = stamped_message


def ping(sc, first):

    global cycle

    if cycle == count - 1 or shut_down:
        pubnub.publish().channel("hello_world").message("hello.py terminated...").pn_async(my_publish_callback)
        pubnub.unsubscribe().channels('hello_world').execute()
        sys.exit(0)
    elif first:
        pubnub.publish().channel("hello_world").message({
        'ping': 'hello.py'
        }).pn_async(my_publish_callback)
        s.enter(interval, 1, ping, (sc, False))
    else:
        for k in data:
            data[k].append("empty")
        cycle += 1
        pubnub.publish().channel("hello_world").message({
            'ping': 'hello.py'
        }).pn_async(my_publish_callback)
        s.enter(interval, 1, ping, (sc, False))


def exit_status():

    t = datetime.now(pytz.timezone('US/Central')).strftime('%Y.%m.%d %H-%M-%S')

    print("Program terminated at: " + t)

    file = open(f"{t} output.txt", "w+")

    for e in allEvents:
        file.write(e)
        file.write("\n")

    file.close()

    workbook = xlsxwriter.Workbook(f"{t}.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 22)
    worksheet.set_column('B:B', 22)
    worksheet.set_column('C:C', 12)
    worksheet.set_column('D:D', 42)
    worksheet.set_column('E:E', 62)
    col = 0

    for k in data:
        row = 0
        worksheet.write(row, col, k)
        row += 1
        for v in data[k]:
            worksheet.write(row, col, v)
            row += 1
        col += 1
    workbook.close()


s.enter(10, 1, ping, (s, True))
atexit.register(exit_status)
pubnub.add_listener(MySubscribeCallback())
pubnub.subscribe().channels('hello_world').execute()
s.run()


