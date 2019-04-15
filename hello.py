from pubnub.callbacks import SubscribeCallback
from pubnub.enums import PNStatusCategory
from pubnub.pnconfiguration import PNConfiguration
from pubnub.pubnub import PubNub
import sched, time
import atexit
import sys
import xlsxwriter
from time import gmtime, strftime
from datetime import datetime



workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Ping')
worksheet.write('B1', 'Ack')
worksheet.write('C1', 'Arduino Time')
worksheet.write('D1', 'RAW PING')
worksheet.write('E1', 'RAW ACK')

pnconfig = PNConfiguration()

pnconfig.subscribe_key = 'sub-c-05dce56c-3c2e-11e7-847e-02ee2ddab7fe'
pnconfig.publish_key = 'pub-c-b6db3020-95a8-4c60-8d16-13345aaf8709'

pubnub = PubNub(pnconfig)

s = sched.scheduler(time.time, time.sleep)

pubnubEvents = []

count = 0
shut_down = False

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
        if status.category == PNStatusCategory.PNUnexpectedDisconnectCategory:
            pass  # This event happens when radio / connectivity is lost

        elif status.category == PNStatusCategory.PNConnectedCategory:
            # Connect event. You can do stuff like publish, and know you'll get it.
            # Or just use the connected event to confirm you are subscribed for
            # UI / internal notifications, etc
            pubnub.publish().channel("hello_world").message("Python script hello.py on channel hello_world...").pn_async(my_publish_callback)
        elif status.category == PNStatusCategory.PNReconnectedCategory:
            pass
            # Happens as part of our regular operation. This event happens when
            # radio / connectivity is lost, then regained.
        elif status.category == PNStatusCategory.PNDecryptionErrorCategory:
            pass
            # Handle message decryption error. Probably client configured to
            # encrypt messages and on live data feed it received plain text.

    def message(self, pubnub, message):
        pass  # Handle new message stored in message.message
        print(message.message)
        t = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        handle_incoming(str(message.message), t)


def handle_incoming(raw_message, timestamp):
    pubnubEvents.append("at: " + timestamp + "    " + raw_message)
    if str(raw_message[2]) == 'x':
        global shut_down
        shut_down = True
    if str(raw_message[2]) == 's':
        if str(raw_message[12:19]) == "pingAck":
            curTime = raw_message[raw_message.find('T')+1:raw_message.find('}')-1]
            print("Current Arduino uptime: " + curTime)
            worksheet.write(count, 1, timestamp)
            worksheet.write(count, 2, curTime)
            worksheet.write(count, 4, "at: " + timestamp + "    " + raw_message)
    if str(raw_message[2]) == 'p':
        if str(raw_message[10:18]) == "hello.py":
            worksheet.write(count, 0, timestamp)
            worksheet.write(count, 3, "at: " + timestamp + "    " + raw_message)




def ping(sc):
    pubnub.publish().channel("hello_world").message({
    'ping': 'hello.py'
    }).pn_async(my_publish_callback)
    s.enter(180, 1, ping, (sc,))

    global count
    global shut_down
    count += 1

    if count == 160 or shut_down:
        pubnub.unsubscribe_all()
        sys.exit(0)


def exit_status(subscribeEvents):
    print("Program terminated")
    file = open("output.txt", "w")


    for e in subscribeEvents:
        file.write(e)
        file.write("\n")

    file.close()
    workbook.close()


s.enter(10, 1, ping, (s,))
atexit.register(exit_status, pubnubEvents)
pubnub.add_listener(MySubscribeCallback())
pubnub.subscribe().channels('hello_world').execute()
s.run()


