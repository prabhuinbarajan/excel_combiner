import uuid
from flask import Flask, request, jsonify, make_response
from flask_restx import Resource, Api, reqparse, fields
from post_process_utils import PostProcessTask, MQPublisher

app = Flask(__name__)
api = Api(app)

app = Flask(__name__)
api = Api(app)
mq = MQPublisher.getInstance()


resource_fields = api.model('PostProcessingTask', {
    'id': fields.String,
    'filepath': fields.String,
    'period': fields.String,
    'year': fields.Integer,
})

# PostProcessList
# shows a list of all post_processs, and lets you POST to add new tasks
@api.route('/post_processs')
class PostProcessList(Resource):
    @api.expect(resource_fields, validate=True)
    @api.doc(body=PostProcessTask)
    def post(self):
        json_data = request.json
        postProcessingTask = PostProcessTask.getPostProcessTaskFromJson(json_data)
        mq.publish(postProcessingTask)
        return make_response(jsonify(postProcessingTask.__dict__), 201)


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True,use_reloader=False)