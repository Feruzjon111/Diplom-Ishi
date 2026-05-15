from django.db import migrations, models


def ensure_balance_column(apps, schema_editor):
    table_name = "app_excel_profile"
    with schema_editor.connection.cursor() as cursor:
        columns = {
            column.name
            for column in schema_editor.connection.introspection.get_table_description(cursor, table_name)
        }

    if "balance" not in columns:
        schema_editor.execute(
            "ALTER TABLE app_excel_profile ADD COLUMN balance decimal NOT NULL DEFAULT 0"
        )

    schema_editor.execute("UPDATE app_excel_profile SET balance = 0 WHERE balance IS NULL")


class Migration(migrations.Migration):

    dependencies = [
        ("app_excel", "0005_student_department_head_student_faculty_dean"),
    ]

    operations = [
        migrations.SeparateDatabaseAndState(
            database_operations=[
                migrations.RunPython(ensure_balance_column, migrations.RunPython.noop),
            ],
            state_operations=[
                migrations.AddField(
                    model_name="profile",
                    name="balance",
                    field=models.DecimalField(decimal_places=2, default=0, max_digits=12),
                ),
            ],
        ),
    ]
