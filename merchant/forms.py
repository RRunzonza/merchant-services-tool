from django import forms


class UploadForm(forms.Form):
    zwg_file = forms.FileField(
        label='Upload ZWG Transactions',
        help_text='Excel file (.xlsx or .xls)',
    )
    usd_file = forms.FileField(
        label='Upload USD Transactions',
        help_text='Excel file (.xlsx or .xls)',
    )

    def clean_zwg_file(self):
        f = self.cleaned_data['zwg_file']
        if not f.name.lower().endswith(('.xlsx', '.xls')):
            raise forms.ValidationError('Please upload a valid Excel file (.xlsx or .xls).')
        return f

    def clean_usd_file(self):
        f = self.cleaned_data['usd_file']
        if not f.name.lower().endswith(('.xlsx', '.xls')):
            raise forms.ValidationError('Please upload a valid Excel file (.xlsx or .xls).')
        return f


class SheetSelectForm(forms.Form):
    zwg_sheet = forms.ChoiceField(label='ZWG Sheet')
    usd_sheet = forms.ChoiceField(label='USD Sheet')

    def __init__(self, *args, zwg_sheets=None, usd_sheets=None, **kwargs):
        super().__init__(*args, **kwargs)
        if zwg_sheets:
            self.fields['zwg_sheet'].choices = [(s, s) for s in zwg_sheets]
        if usd_sheets:
            self.fields['usd_sheet'].choices = [(s, s) for s in usd_sheets]


class CIFLookupForm(forms.Form):
    cif = forms.CharField(
        label='Customer ID',
        max_length=6,
        widget=forms.TextInput(attrs={'placeholder': 'e.g. 012499', 'class': 'form-control'}),
    )

    def clean_cif(self):
        val = self.cleaned_data['cif'].strip()
        if '.' in val:
            raise forms.ValidationError('Customer ID must be a whole number, not a decimal.')
        if val.startswith('-'):
            raise forms.ValidationError('Customer ID must be a positive number.')
        if not val.isdigit():
            raise forms.ValidationError('Customer ID must contain digits only.')
        if len(val) != 6:
            direction = 'too short' if len(val) < 6 else 'too long'
            raise forms.ValidationError(
                f'Customer ID is {direction} ({len(val)} digits). It must be exactly 6 digits.'
            )
        return val


class DateRangeForm(forms.Form):
    from_date = forms.DateField(
        label='From Date',
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
    )
    to_date = forms.DateField(
        label='To Date',
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
    )

    def clean(self):
        cleaned = super().clean()
        from_d = cleaned.get('from_date')
        to_d   = cleaned.get('to_date')
        if from_d and to_d and from_d > to_d:
            raise forms.ValidationError('FROM date cannot be after TO date.')
        return cleaned
