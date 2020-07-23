import zmq
from post_process_utils import _topic

port = "5556"
context = zmq.Context()
socket = context.socket(zmq.SUB)
socket.connect("tcp://localhost:%s" % port)
socket.setsockopt(zmq.SUBSCRIBE, _topic.encode('utf-8'))
print('waiting for message')
topic = socket.recv_string()
md = socket.recv_json()
print(md)
