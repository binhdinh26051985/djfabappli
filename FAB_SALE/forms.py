from django.forms import ModelForm
from .models import *

class zohform(ModelForm):
    class Meta:
        model = zoh
        fields = '__all__'
        
class millform(ModelForm):
    class Meta:
        model = mill
        fields = '__all__'
        
class factoryform(ModelForm):
    class Meta:
        model = factory
        fields = '__all__'
        
                
class fabPOform(ModelForm):
    class Meta:
        model = fabPO
        fields = '__all__'
    def __init__(self, *args, **kwargs):
      super(fabPOform, self).__init__(*args, **kwargs)
      self.user = kwargs.pop('user', None)
      self.fields['Zoh'].queryset = self.fields['Zoh'].queryset.order_by('ZOH')
      self.fields['Mill'].queryset = self.fields['Mill'].queryset.order_by('MILL')
      self.fields['Ship_to'].queryset = self.fields['Ship_to'].queryset.order_by('SHIP_to')