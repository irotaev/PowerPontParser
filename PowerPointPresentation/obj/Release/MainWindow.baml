   M S B A M L   `   `   `  ���� U  QMahApps.Metro, Version=0.12.1.0, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls   PowerPointPresentation'   !PowerPointPresentation.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     PControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro     -   $��	AppWindow��  � Title$+  %Парсинг презентации��$	��320�� � SizeToContent$ Height=� � 
ResizeMode$ NoResize=� � Icon$ /Resource/PowerPoint.ico��   
FlatWindow#��[   � 	Resources   StoryboardMessagePopUp �� 
TargetName �� TargetProperty W� From	 W� To	 g� Duration
 g� 	BeginTime��%�   &	       �� .��W� $ MessagePopUp��$ Opacity�$	 0.0X�$	 1.0X�$	 0:0:0:1D�$
 0:0:0X�W� $ MessagePopUp��$ Opacity�$	 1.0X�$	 0.0X�$	 0:0:1D�$
 0:0:4X�.��� -   $��GridWrapper��.+��� -   $��MessagePopUp��$�� ��$��Center=�$	��Top=�$��10q� �� Padding$ 5q�$��16(� �� 
LineHeight$ 30��
��� � � K� ZIndex$	 100��$��30��
������� J� Opacity$ 0Q��� -   $��ProgressInfo��$��Center=�$��Center=�$��10q�$��16(�
��� � �$
��Wrap=�$y�	Collapsed=�� -   $��ProgressBar��$��Center=�$��Bottom=�$	��150��$y�	Collapsed=�$��0 0 0 20q�$��15��� -   $��MainGrid��$��3q� � RowDefinitions �� $��* ��� $��* ��� $��* ��� $��* ��� $��* ��� $��* �.+��� $��0������ .��&$Файл с презентацией.���  � ColumnDefinitions �� $��2* ��� $��1* �.+��� -   $��PresentationFileName��$��5q�$��0��$
��Wrap=�$��15���� -   $��0 5 5 5q�$��1��$��Выбрать���� $��1������ .��)'Название презентации.���� -	   $��PresentationName��$��50�� �� AcceptsReturn .  �� TextWrapping$
 Wrap=�$��5 0 5 0q��� $��2������ .��+)Заголовок презентации.���� -
   $��PresentationTitle��$��50�� . $
 Wrap=�$��5 0 5 0q��� $��3������ .��Login.���� -   $��Login��$��50�� . $
 Wrap=�$��5 0 5 0q��� $��3��$y�	Collapsed=����� .��)'Url news (Необязательное).���� -   $��UrlNews��$��50�� . $
 Wrap=�$��5 0 5 0q��� $��4��$y�	Collapsed=����� .��+)Категория презентации.���� -   $��CategorieComboBox�� �� SelectedValuePath$	 Key�� �� DisplayMemberPath$ Value��$��5q� �� SelectedIndex$ 0������  �� RelativeSource ��  �� AncestorType$ WindowL� �� Mode$ FindAncestor=� �� Path$ 
Categories�]� $��5��$��5q���. .+��� -   $��30��$��10 0 0 0q�$��30��$��Right=�$y�	Collapsed=�.���� $��/Resource/settings.png���� -   $1��+Распарсить презентацию��$��30��$��Stretch=�