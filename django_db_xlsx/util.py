import numbers

from django.apps import apps
from django.conf import settings
from django.http import HttpResponse
from django.utils.encoding import smart_str

from openpyxl import Workbook


def get_model_export_headers(model):
    """ Get header names, friendly for exporting data to csv/xlsx """
    for field in model._meta.get_fields():
        if 'ManyToManyRel' != field.__class__.__name__ and 'ManyToOneRel' != field.__class__.__name__:
            yield field.name


def get_model_export_row(model, obj, many_to_many_delimeter=' '):
    """ Get a single model objects row data for export """
    for field in model._meta.get_fields():
        if 'ManyToManyRel' == field.__class__.__name__:
            continue
        if 'ManyToOneRel' == field.__class__.__name__:
            continue
        if 'ManyToManyField' == field.__class__.__name__:
            pks = [str(i.pk) for i in getattr(obj, field.name).all()]
            yield many_to_many_delimeter.join(pks)
        else:
            attribute = getattr(obj, field.name)
            if attribute is None:
                yield None
            elif isinstance(attribute, numbers.Number):
                yield attribute
            elif hasattr(attribute, 'pk'):
                yield attribute.pk
            else:
                yield smart_str(attribute)


def load_models(wb, target_models=None):
    """ Update all ourdesign models with data in the workbook """
    if target_models is None:
        target_models = settings.DJANGO_DB_XLSX_MODELS
    for app_label, model_name in target_models:
        model = apps.get_model(app_label, model_name)
        ws = wb.get_sheet_by_name(name=model_name)
        # TODO delete what isnt in ws?
        headers = [cell.value for cell in list(ws.rows)[0]]
        for i, row in enumerate(ws.rows):
            if i == 0:
                continue
            row_values = [cell.value for cell in row]
            kwargs = dict(zip(headers, row_values))
            model_pk = kwargs['id']
            del kwargs['id']
            post_process_m2m = {}
            for header in headers:
                if header in kwargs and kwargs[header] is None:
                    del kwargs[header]
                else:
                    # update value in kwargs to expected field type
                    field = model._meta.get_field(header)
                    if field.__class__.__name__ == 'ManyToManyField':
                        try:
                            ids = [int(model_id) for model_id in kwargs[header].split(',')]
                            objects = field.related_model.objects.filter(id__in=ids)
                            post_process_m2m[header] = objects
                        except ValueError:
                            print('ValueError, m2m object likely empty: {}'.format(header))
                        del kwargs[header]
                    elif field.__class__.__name__ == 'ForeignKey':
                        field_pk = kwargs[header]
                        try:
                            kwargs[header] = field.related_model.objects.get(pk=field_pk)
                        except:
                            template = "Model: {} field: {} FK doesn't exist with id: {}"
                            print(template.format(model_name, header, field_pk))
                            del kwargs[header]
            model_object, created = model.objects.update_or_create(pk=model_pk, defaults=kwargs)
            for field_name, value in post_process_m2m.items():
                field = model._meta.get_field(field_name)
                m2m_manager = getattr(model_object, field.name)
                m2m_manager.clear()
                m2m_manager.add(*value)
            print('Created' if created else 'Updated' + ' object from spreadsheet: ' +
                  model_name + ' ' + str(model_pk))


def dump_models(target_models=None, path=None, wb=None):
    """ Dump models to excel response or file path if given """
    if target_models is None:
        target_models = settings.DJANGO_DB_XLSX_MODELS
    if not wb:
        wb = Workbook()
    ws1 = wb.active
    # models to save
    for app_label, model_name in target_models:
        model = apps.get_model(app_label, model_name)
        ws = wb.create_sheet(title=model_name)
        ws.append(get_model_export_headers(model))
        for obj in model.objects.all():
            column_data = get_model_export_row(model, obj, many_to_many_delimeter=',')
            row_data = ['' if column_datum is None else column_datum for column_datum in column_data]
            ws.append(row_data)
    if 'Sheet' == ws1.title:  # delete default page
        wb.remove_sheet(ws1)
    if path:
        wb.save(path)
    else:
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename=ourdesign.xlsx'
        wb.save(response)
        return response
