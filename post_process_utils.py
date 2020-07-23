import json
import pickle
import zmq

port = "5556"
_topic = 'postprocess'


class MQPublisher:
    socket_init = False
    context = None
    pub_socket = None
    __instance = None

    @staticmethod
    def getInstance():
        """ Static access method. """
        if MQPublisher.__instance == None:
            MQPublisher()
        return MQPublisher.__instance

    def __init__(self):
        """ Virtually private constructor. """
        if MQPublisher.__instance != None:
            raise Exception("This class is a singleton!")
        else:
            self.context = zmq.Context()
            self.pub_socket = self.context.socket(zmq.PUB)
            self.pub_socket.bind("tcp://*:%s" % port)
            MQPublisher.__instance = self

    def publish(self, obj, protocol=-1):
        #p = pickle.dumps(obj, protocol)
        self.pub_socket.send_string(_topic, zmq.SNDMORE)
        self.pub_socket.send_json(obj.__dict__)
        return


class PostProcessTask:

    def __init__(self, *args, **kwargs):
        self.id = kwargs['id']
        self.filepath = kwargs['filepath']
        self.period = kwargs['period']
        self.year = kwargs['year']

    @classmethod
    def getPostProcessTask(cls, payload):
        postProcessTask = PostProcessTask(**json.loads(payload))
        return postProcessTask

    @classmethod
    def getPostProcessTaskFromJson(cls, jsonObj):
        postProcessTask = PostProcessTask(**jsonObj)
        return postProcessTask

    @classmethod
    def getPostProcessPayload(cls, obj):
        payload = json.dumps(obj.__dict__, indent=4)
        return payload

# pl = '{"id": "'+ str(uuid.uuid4()) + '", "filepath": "path1", "period": "P6", "year": 2020}'
# ppObj = PostProcessTask.getPostProcessTask(pl)
# ppStr = PostProcessTask.getPostProcessPayload(ppObj)
# print(ppStr)


#import time

#from pynng import Push0, Pull0, Timeout
#addr = 'tcp://127.0.0.1:31313'
#push= Push0(listen=addr)
#time.sleep(0.01)
