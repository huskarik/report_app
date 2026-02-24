from flask_wtf import FlaskForm
from wtforms import SelectField, DateTimeLocalField, SubmitField
from wtforms.validators import DataRequired, ValidationError


class PeriodForm(FlaskForm):
    projects = SelectField(
        'Выберите значение',
        choices=[
        ],
        validators=[DataRequired()]
    )

    date_from = DateTimeLocalField('Дата ОТ', format='%Y-%m-%dT%H:%M', validators=[DataRequired()])
    date_to = DateTimeLocalField('Дата ДО', format='%Y-%m-%dT%H:%M', validators=[DataRequired()])

    submit = SubmitField('Отправить')


def validate_date_to(form, field):
    if form.date_from.data and field.data:
        if field.data < form.date_from.data:
            raise ValidationError('Дата "ДО" не может быть раньше даты "ОТ"')