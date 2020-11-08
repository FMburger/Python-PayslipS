import os
from flask import Flask, render_template, request, flash, url_for, redirect
from flask_bootstrap import Bootstrap
from forms import LoginForm
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import login_user, logout_user, login_required,  LoginManager,  current_user, UserMixin
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
import json
import configparser

# config.init
config = configparser.ConfigParser()
config.read('config.ini')

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config['SECRET_KEY'] =\
    '\x0f[kH\xf0*n\xcao\xb08\x86\x08$\xcf\xf5P>\xf7\xb5\xb9\xb0\x1c\xa4'
app.config['SQLALCHEMY_DATABASE_URI'] =\
    'sqlite:///' + os.path.join(basedir, 'user.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

bootstrap = Bootstrap(app)
db = SQLAlchemy(app)
migrate = Migrate(app, db)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'


class Role(db.Model):
    __tablename__ = 'roles'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), unique=True)
    users = db.relationship('User', backref='role', lazy='dynamic')

    def __repr__(self):
        return '<Role %r>' % self.name


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(64), unique=True, index=True)
    username = db.Column(db.String(64), unique=True, index=True)
    role_id = db.Column(db.Integer, db.ForeignKey('roles.id'))
    password_hash = db.Column(db.String(128))

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    @property
    def password(self):
        raise AttributeError('password is not a readable attribute')

    @password.setter
    def password(self, password):
        self.password_hash = generate_password_hash(password)

    def verify_password(self, password):
        return check_password_hash(self.password_hash, password)


@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404


@app.errorhandler(500)
def internal_server_error(e):
    return render_template('500.html'), 500


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(email=form.email.data.lower()).first()
        if user is not None and user.verify_password(form.password.data):
            login_user(user, form.remember_me.data)
            next = request.args.get('next')
            if next is None or not next.startswith('/'):
                next = url_for('index')
            return redirect(next)
        flash('Invalid email or password.')
    return render_template('login.html', form=form)


@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out.')
    return redirect(url_for('index'))


@app.route('/sender', methods=['GET', 'POST'])
@login_required
def sender():
    import controller
    controller = controller.Controller()

    if request.method == 'POST':
        obj = request.get_json()
        if obj['num'] == 1:
            result = controller.model_employeeList.get_list_departments(
                obj['payPeriod']
            )
        elif obj['num'] == 2:
            result = controller.model_employeeList.get_list_employees(
                obj['payPeriod'],
                obj['department']
            )
        elif obj['num'] == 3:
            result = controller.email_payslip(obj['payPeriod'], obj['department'], obj['employee'])
        response = app.response_class(
            response=json.dumps(result),
            status=200,
            mimetype='application/json'
        )
        return response
    else:
        # default value
        payPeriods = controller.default_payPeriod_list
        departments = controller.default_department_list
        employees = controller.default_employee_list
        return render_template(
            'sender.html',
            payPeriods=payPeriods,
            departments=departments,
            employees=employees
        )


@app.route('/myPayslip')
@login_required
def my_payslip():
    return render_template('my_payslip.html')


@app.route('/setting', methods=['GET', 'POST'])
@login_required
def setting():
    if request.method == 'POST':
        obj = request.get_json()
        config['smtp']['smtp_host'] = obj['smtpServer']
        config['smtp']['smtp_port'] = obj['port']
        config['smtp']['user_name'] = obj['id']
        config['smtp']['user_password'] = obj['password']
        config['email']['email_content'] = obj['email_content']
        with open('config.ini', 'w') as configfile:
            config.write(configfile)
        result = '儲存成功'
        response = app.response_class(
            response=json.dumps(result),
            status=200,
            mimetype='application/json'
        )
        return response
    else:
        return render_template(
            'setting.html',
            server=config['smtp']['smtp_host'],
            port=config['smtp']['smtp_port'],
            id=config['smtp']['user_name'],
            password=config['smtp']['user_password'],
            emailContent=config['email']['email_content']
        )


if __name__ == '__main__':
    app.run()
