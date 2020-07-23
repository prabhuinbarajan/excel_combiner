import uuid

from flask import Flask, request, jsonify, make_response
from flask_restx import Resource, Api, reqparse, fields



from post_process_utils import PostProcessTask, MQPublisher

app = Flask(__name__)
api = Api(app)

app = Flask(__name__)
api = Api(app)
mq = MQPublisher.getInstance()

TODOS = {
    'post_process1': {'id': str(uuid.uuid4()), 'filepath': 'path1', 'period': 'P6', 'year': 2020},
    'post_process2': {'id': str(uuid.uuid4()), 'filepath': 'path2', 'period': 'P7', 'year': 2020},
    'post_process3': {'id': str(uuid.uuid4()), 'filepath': 'path3', 'period': 'P6', 'year': 2020},
}

resource_fields = api.model('PostProcessingTask', {
    'id': fields.String,
    'filepath': fields.String,
    'period': fields.String,
    'year': fields.Integer,
})



def abort_if_post_process_doesnt_exist(post_process_id):
    if post_process_id not in TODOS:
        abort(404, message="PostProcess {} doesn't exist".format(post_process_id))

parser = reqparse.RequestParser()
parser.add_argument('task')

"""
# PostProcess
# shows a single post_process item and lets you delete a post_process item
@api.route('/post_processs/<string:post_process_id>')
class PostProcess(Resource):
    def get(self, post_process_id):
        abort_if_post_process_doesnt_exist(post_process_id)
        return TODOS[post_process_id]

    def delete(self, post_process_id):
        abort_if_post_process_doesnt_exist(post_process_id)
        del TODOS[post_process_id]
        return '', 204

    def put(self, post_process_id):
        args = parser.parse_args()
        task = {'task': args['task']}
        TODOS[post_process_id] = task
        return task, 201

"""

# PostProcessList
# shows a list of all post_processs, and lets you POST to add new tasks
@api.route('/post_processs')
class PostProcessList(Resource):
    def get(self):
        return TODOS

    @api.expect(resource_fields, validate=True)
    @api.doc(body=PostProcessTask)
    def post(self):
        json_data = request.json
        postProcessingTask = PostProcessTask.getPostProcessTaskFromJson(json_data)
        mq.publish(postProcessingTask)
        return make_response(jsonify(postProcessingTask.__dict__), 201)


if __name__ == '__main__':
    app.run(debug=True,use_reloader=False)