# Generated by Django 3.2.5 on 2023-06-21 09:25

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='EmailModel',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('Subject', models.TextField()),
                ('Sender', models.TextField()),
                ('Received_Time', models.TextField()),
                ('Email_content', models.TextField()),
            ],
        ),
    ]
