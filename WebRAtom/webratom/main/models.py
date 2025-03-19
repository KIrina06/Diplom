# This is an auto-generated Django model module.
# You'll have to do the following manually to clean this up:
#   * Rearrange models' order
#   * Make sure each model has one field with primary_key=True
#   * Make sure each ForeignKey and OneToOneField has `on_delete` set to the desired behavior
#   * Remove `managed = False` lines if you wish to allow Django to create, modify, and delete the table
# Feel free to rename the models, but don't rename db_table values or field names.
from django.db import models


class AuthGroup(models.Model):
    name = models.CharField(unique=True, max_length=150)

    class Meta:
        managed = False
        db_table = 'auth_group'


class AuthGroupPermissions(models.Model):
    id = models.BigAutoField(primary_key=True)
    group = models.ForeignKey(AuthGroup, models.DO_NOTHING)
    permission = models.ForeignKey('AuthPermission', models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_group_permissions'
        unique_together = (('group', 'permission'),)


class AuthPermission(models.Model):
    name = models.CharField(max_length=255)
    content_type = models.ForeignKey('DjangoContentType', models.DO_NOTHING)
    codename = models.CharField(max_length=100)

    class Meta:
        managed = False
        db_table = 'auth_permission'
        unique_together = (('content_type', 'codename'),)


class AuthUser(models.Model):
    password = models.CharField(max_length=128)
    last_login = models.DateTimeField(blank=True, null=True)
    is_superuser = models.BooleanField()
    username = models.CharField(unique=True, max_length=150)
    first_name = models.CharField(max_length=150)
    last_name = models.CharField(max_length=150)
    email = models.CharField(max_length=254)
    is_staff = models.BooleanField()
    is_active = models.BooleanField()
    date_joined = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'auth_user'


class AuthUserGroups(models.Model):
    id = models.BigAutoField(primary_key=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)
    group = models.ForeignKey(AuthGroup, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_user_groups'
        unique_together = (('user', 'group'),)


class AuthUserUserPermissions(models.Model):
    id = models.BigAutoField(primary_key=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)
    permission = models.ForeignKey(AuthPermission, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_user_user_permissions'
        unique_together = (('user', 'permission'),)


class BudgetingType(models.Model):
    id_budgeting_type = models.AutoField(primary_key=True)
    name_budgeting_type = models.TextField()

    class Meta:
        managed = False
        db_table = 'budgeting_type'


class CategoryIndicators(models.Model):
    id_category = models.AutoField(primary_key=True)
    name_category = models.TextField()

    class Meta:
        managed = False
        db_table = 'category_indicators'


class Complex(models.Model):
    id_complex = models.AutoField(primary_key=True)
    full_complex_name = models.TextField()
    short_complex_name = models.TextField(blank=True, null=True)
    area_complex_km2 = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    address_complex = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'complex'


class CostIndicatorsMonthlyVat(models.Model):
    id_indicator = models.AutoField(primary_key=True)
    id_ip = models.ForeignKey('InvestmentProject', models.DO_NOTHING, db_column='id_ip')
    id_object = models.ForeignKey('Object', models.DO_NOTHING, db_column='id_object')
    id_category = models.ForeignKey(CategoryIndicators, models.DO_NOTHING, db_column='id_category')
    id_expense_direction = models.ForeignKey('ExpenseDirection', models.DO_NOTHING, db_column='id_expense_direction')
    price_level = models.CharField(max_length=10, blank=True, null=True)
    construction_works = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    equipment = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    design_works = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    other = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    miscellaneous = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    month = models.CharField(max_length=20, blank=True, null=True)
    year = models.IntegerField(blank=True, null=True)
    fact_or_plan = models.CharField(max_length=10, blank=True, null=True)
    cumulative_effect = models.BooleanField()

    class Meta:
        managed = False
        db_table = 'cost_indicators_monthly_vat'


class Country(models.Model):
    id_country = models.AutoField(primary_key=True)
    name_country = models.CharField(max_length=70)
    code_country = models.CharField(max_length=10, blank=True, null=True)
    capital = models.CharField(max_length=70, blank=True, null=True)
    continent = models.CharField(max_length=20, blank=True, null=True)
    official_language = models.TextField(blank=True, null=True)
    currency = models.TextField(blank=True, null=True)
    time_zone = models.CharField(max_length=50, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'country'


class CuratorIp(models.Model):
    id_curator = models.AutoField(primary_key=True)
    id_ip = models.ForeignKey('InvestmentProject', models.DO_NOTHING, db_column='id_ip')
    id_individual = models.ForeignKey('Individual', models.DO_NOTHING, db_column='id_individual')
    start_date = models.DateField()
    end_date = models.DateField()

    class Meta:
        managed = False
        db_table = 'curator_ip'


class CustomerIp(models.Model):
    id_customer = models.AutoField(primary_key=True)
    id_ip = models.ForeignKey('InvestmentProject', models.DO_NOTHING, db_column='id_ip')
    id_organization = models.ForeignKey('Organization', models.DO_NOTHING, db_column='id_organization')
    participation_start_date = models.DateField()
    participation_end_date = models.DateField()

    class Meta:
        managed = False
        db_table = 'customer_ip'


class Division(models.Model):
    id_division = models.AutoField(primary_key=True)
    name_division = models.CharField(max_length=50)

    class Meta:
        managed = False
        db_table = 'division'


class DjangoAdminLog(models.Model):
    action_time = models.DateTimeField()
    object_id = models.TextField(blank=True, null=True)
    object_repr = models.CharField(max_length=200)
    action_flag = models.SmallIntegerField()
    change_message = models.TextField()
    content_type = models.ForeignKey('DjangoContentType', models.DO_NOTHING, blank=True, null=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'django_admin_log'


class DjangoContentType(models.Model):
    app_label = models.CharField(max_length=100)
    model = models.CharField(max_length=100)

    class Meta:
        managed = False
        db_table = 'django_content_type'
        unique_together = (('app_label', 'model'),)


class DjangoMigrations(models.Model):
    id = models.BigAutoField(primary_key=True)
    app = models.CharField(max_length=255)
    name = models.CharField(max_length=255)
    applied = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'django_migrations'


class DjangoSession(models.Model):
    session_key = models.CharField(primary_key=True, max_length=40)
    session_data = models.TextField()
    expire_date = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'django_session'


class Documents(models.Model):
    id_document = models.AutoField(primary_key=True)
    id_ip = models.ForeignKey('InvestmentProject', models.DO_NOTHING, db_column='id_ip')
    id_object = models.ForeignKey('Object', models.DO_NOTHING, db_column='id_object')
    document_name = models.TextField()
    creation_date = models.DateField()
    last_modified = models.DateTimeField()
    document_type = models.CharField(max_length=20, blank=True, null=True)
    link = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'documents'


class ExpenseDirection(models.Model):
    id_expense_direction = models.AutoField(primary_key=True)
    name_expense_direction = models.TextField()

    class Meta:
        managed = False
        db_table = 'expense_direction'


class Individual(models.Model):
    id_individual = models.AutoField(primary_key=True)
    id_organization = models.ForeignKey('Organization', models.DO_NOTHING, db_column='id_organization')
    last_name = models.CharField(max_length=50)
    first_name = models.CharField(max_length=30)
    middle_name = models.CharField(max_length=30, blank=True, null=True)
    position = models.CharField(max_length=30, blank=True, null=True)
    phone = models.CharField(max_length=15, blank=True, null=True)
    email = models.CharField(max_length=50, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'individual'


class InvestmentProject(models.Model):
    id_ip = models.AutoField(primary_key=True)
    id_country = models.ForeignKey(Country, models.DO_NOTHING, db_column='id_country')
    full_ip_name = models.TextField()
    short_ip_name = models.TextField(blank=True, null=True)
    goal_ip = models.TextField(blank=True, null=True)
    code_ip = models.CharField(max_length=40, blank=True, null=True)
    start_date_ip = models.IntegerField(blank=True, null=True)
    end_date_ip = models.IntegerField(blank=True, null=True)
    decision_level = models.TextField(blank=True, null=True)
    significance_flag = models.BooleanField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'investment_project'


class KeyEvents(models.Model):
    id_key_event = models.AutoField(primary_key=True)
    id_ip = models.ForeignKey(InvestmentProject, models.DO_NOTHING, db_column='id_ip')
    id_object = models.ForeignKey('Object', models.DO_NOTHING, db_column='id_object')
    event_name = models.TextField()
    work_number = models.TextField(blank=True, null=True)
    start_event_date = models.DateField()
    end_event_date = models.DateField()
    event_type = models.CharField(max_length=20, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'key_events'


class MainComplex(models.Model):
    id_complex = models.AutoField(primary_key=True)
    full_complex_name = models.TextField()
    short_complex_name = models.TextField(blank=True, null=True)
    area_complex_km2 = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    address_complex = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'main_complex'


class MainDivision(models.Model):
    id_division = models.AutoField(primary_key=True)
    name_division = models.CharField(max_length=50)

    class Meta:
        managed = False
        db_table = 'main_division'


class MainMymodel(models.Model):
    id = models.BigAutoField(primary_key=True)
    name = models.CharField(max_length=255)
    data = models.TextField()
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'main_mymodel'


class MainRegion(models.Model):
    id_region = models.AutoField(primary_key=True)
    name_region = models.CharField(max_length=100)
    code_region = models.CharField(max_length=50, blank=True, null=True)
    area_region_km2 = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    population = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    building_density = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    urbanization_level = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'main_region'


class Object(models.Model):
    id_object = models.AutoField(primary_key=True)
    id_ip = models.IntegerField(blank=True, null=True)
    id_region = models.ForeignKey('Region', models.DO_NOTHING, db_column='id_region')
    id_division = models.ForeignKey(Division, models.DO_NOTHING, db_column='id_division')
    id_complex = models.ForeignKey(Complex, models.DO_NOTHING, db_column='id_complex', blank=True, null=True)
    id_main_object = models.IntegerField(blank=True, null=True)
    code_by_classifier = models.CharField(max_length=14)
    full_object_name = models.TextField()
    short_object_name = models.TextField(blank=True, null=True)
    visibility_level = models.IntegerField(blank=True, null=True)
    security_classification = models.IntegerField(blank=True, null=True)
    address_object = models.TextField(blank=True, null=True)
    area_object_km2 = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'object'
    
    def __str__(self):
           return self.full_object_name


class ObjectProgram(models.Model):
    id_object_program = models.AutoField(primary_key=True)
    id_program = models.ForeignKey('Program', models.DO_NOTHING, db_column='id_program')
    id_object = models.ForeignKey(Object, models.DO_NOTHING, db_column='id_object')

    class Meta:
        managed = False
        db_table = 'object_program'


class ObjectStep(models.Model):
    id_object_step = models.AutoField(primary_key=True)
    id_step = models.ForeignKey('Step', models.DO_NOTHING, db_column='id_step')
    id_object = models.ForeignKey(Object, models.DO_NOTHING, db_column='id_object')
    start_date = models.DateField()
    end_date = models.DateField()

    class Meta:
        managed = False
        db_table = 'object_step'


class Organization(models.Model):
    id_organization = models.AutoField(primary_key=True)
    full_name = models.TextField()
    short_name = models.TextField(blank=True, null=True)
    okved = models.CharField(max_length=20, blank=True, null=True)
    okpo = models.CharField(max_length=20, blank=True, null=True)
    duns = models.CharField(max_length=20, blank=True, null=True)
    address = models.TextField(blank=True, null=True)
    contact_phone = models.CharField(max_length=15, blank=True, null=True)
    email = models.CharField(max_length=50, blank=True, null=True)
    website = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'organization'


class Phase(models.Model):
    id_phase = models.AutoField(primary_key=True)
    phase_number = models.IntegerField()
    full_phase_name = models.TextField()
    short_phase_name = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'phase'


class Program(models.Model):
    id_program = models.AutoField(primary_key=True)
    id_budgeting_type = models.ForeignKey(BudgetingType, models.DO_NOTHING, db_column='id_budgeting_type')
    full_program_name = models.TextField()
    short_program_name = models.TextField(blank=True, null=True)
    description_program = models.TextField(blank=True, null=True)
    document_program = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'program'


class ProjectKvl(models.Model):
    id_pr_kvl = models.AutoField(primary_key=True)
    id_ip = models.ForeignKey(InvestmentProject, models.DO_NOTHING, db_column='id_ip')
    name_pr_kvl = models.TextField()
    goal_pr_kvl = models.TextField(blank=True, null=True)
    capital_investment_character = models.TextField(blank=True, null=True)
    technical_specification_approved = models.BooleanField()
    docs_approved = models.BooleanField()
    year = models.IntegerField(blank=True, null=True)
    fact_or_plan = models.CharField(max_length=10, blank=True, null=True)
    cumulative_effect = models.BooleanField()

    class Meta:
        managed = False
        db_table = 'project_kvl'


class Region(models.Model):
    id_region = models.AutoField(primary_key=True)
    name_region = models.CharField(max_length=100)
    code_region = models.CharField(max_length=50, blank=True, null=True)
    area_region_km2 = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    population = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    building_density = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)
    urbanization_level = models.DecimalField(max_digits=65535, decimal_places=65535, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'region'


class Stage(models.Model):
    id_stage = models.AutoField(primary_key=True)
    id_phase = models.ForeignKey(Phase, models.DO_NOTHING, db_column='id_phase')
    stage_number = models.IntegerField()
    full_stage_name = models.TextField()
    short_stage_name = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'stage'


class Step(models.Model):
    id_step = models.AutoField(primary_key=True)
    id_stage = models.ForeignKey(Stage, models.DO_NOTHING, db_column='id_stage', blank=True, null=True)
    step_number = models.IntegerField()
    full_step_name = models.TextField()
    short_step_name = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'step'


class TechCustomer(models.Model):
    id_tech_customer = models.AutoField(primary_key=True)
    id_object = models.ForeignKey(Object, models.DO_NOTHING, db_column='id_object')
    id_organization = models.ForeignKey(Organization, models.DO_NOTHING, db_column='id_organization')
    participation_start_date = models.DateField()
    participation_end_date = models.DateField()

    class Meta:
        managed = False
        db_table = 'tech_customer'
