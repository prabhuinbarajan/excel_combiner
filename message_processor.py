import time

from zmq.utils import jsonapi

"""
from pynng import Push0, Pull0, Timeout

addr = 'tcp://127.0.0.1:31313'
with Push0(listen=addr) as push, \
        Pull0(dial=addr, recv_timeout=100) as pull0:
    pass
    # give some time to connect
    time.sleep(0.01)
    push.send(b'hi some node')
    push.send(b'hi some other node')
    print(pull0.recv())  # prints b'hi some node'
    print(pull0.recv())  # prints b'hi some node'

    try:
        print('msg {}'.format(pull0.recv()))  # prints b'hi some node'
        assert False, "Cannot get here, since messages are sent round robin"
    except Timeout:
        print ('here')
        pass


import sys
import zmq

port = "5556"
context = zmq.Context()
socket = context.socket(zmq.SUB)
socket.connect("tcp://localhost:%s" % port)

print('waiting for message')
md = socket.recv_json()
print(md)

"""

#!/usr/bin/env python
# coding: utf8
"""
Experiments with 0MQ PUB/SUB pattern.
Creates a publisher with 26 topics (A, B, ... Z) and
spawns clients that randomly subscribe to a subset
of the available topics. Console output shows 
who subscribed to what, when topic updates are sent
and when clients receive the messages.
Runs until killed.
Author: Michael Ellis
License: WTFPL
"""
import os
import sys
import time
import zmq
from multiprocessing import Process
from random import sample, choice
import json

PORT = 5566
TOPICS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" # split into ['A', 'B', ... ]
PUBSLEEP = 10 # Sleep time at bottom of publisher() loop.
NCLIENTS = 1  # Number of clients spawned.
NSUBS = 3  # Number of topics each client subscribes to.

assert NSUBS <= len(TOPICS)

def mogrify(topic, msg):
    """ json encode the message and prepend the topic """
    return topic + ' ' + json.dumps(msg)

def demogrify(topicmsg):
    """ Inverse of mogrify() """
    json0 = topicmsg.find('{')
    topic = topicmsg[0:json0].strip()
    msg = json.loads(topicmsg[json0:])
    return topic, msg
json = {
    "test": "test",
    "key": "value"
}
topic = 'camera_frame'

def publisher():
    """ Randomly update and publish topics """
    context = zmq.Context()
    sock = context.socket(zmq.PUB)
    sock.bind("tcp://*:{}".format(PORT))


    while True:
        try:
            sock.send_string(topic, zmq.SNDMORE)
            sock.send_json(json)
            print ("Sent topic {}".format(json))
            time.sleep(PUBSLEEP)
        except KeyboardInterrupt:
            sys.exit()

def client(number, topics):
    """
    Subscribe to list of topics and wait for messages.
    """
    context = zmq.Context()
    sock = context.socket(zmq.SUB)
    sock.connect("tcp://localhost:{}".format(PORT))
    print("client connected")
    sock.setsockopt(zmq.SUBSCRIBE, topic.encode('utf_8'))

    while True:
        try:
            #topic, msg = demogrify(sock.recv())
            val = sock.recv_string()
            msg = sock.recv_json()
            print("Client{}  {}".format(val, msg))
            sys.stdout.flush()
        except KeyboardInterrupt:
            sys.exit()

_procd = dict()
def run():
    """ Spawn publisher and clients. Loop until terminated. """
    ## Launch publisher
    name = 'publisher'
    _procd[name] = Process(target=publisher, name=name)
    _procd[name].start()

    ## Launch the subscribers
    for n in range(NCLIENTS):
        name = 'client{}'.format(n)
        _procd[name] = Process(target=client,
                               name=name,
                               args=(n, sample(TOPICS, NSUBS)))
        _procd[name].start()


    ## Sleep until killed
    while True:
        time.sleep(1)

if __name__ == '__main__':
    import signal
    def handler(signum, frame):
        """ Handler for SIGTERM """
        # kill the processes we've launched
        try:
            for _, proc in _procd.iteritems():
                if proc and proc.is_alive():
                    proc.terminate()
        finally:
            sys.exit()

    signal.signal(signal.SIGTERM, handler)

    run()