   M S B A M L   `   `   `  ����U  QMahApps.Metro, Version=0.12.1.0, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  5   	    PowerPointPresentation'   !PowerPointPresentation.MainWindow   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   PControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro     5   	   -   5        $��	AppWindow��5
   	     � Title$+  %Парсинг презентации��5   	   $	��320��5   	    � SizeToContent$ Height=�5   	    � 
ResizeMode$ NoResize=�5   	    � Icon$ /Resource/PowerPoint.ico��5	   	      
FlatWindow#��[  5
       � 	Resources 5        StoryboardMessagePopUp �� 
TargetName �� TargetProperty W� From	 W� To	 g� Duration
 g� 	BeginTime��%  &	       �� 5      .��5      W� $ MessagePopUp��6   $ Opacity�6=   $	 0.0X�6a   $	 1.0X�6l   $	 0:0:0:1D�6u   $
 0:0:0X�6�   W� 5      $ MessagePopUp��6   $ Opacity�6=   $	 1.0X�6a   $	 0.0X�6l   $	 0:0:1D�6u   $
 0:0:4X�6�   5      5      .��5      � .+�5      �� -   5        $��MessagePopUp��5      $�� ��6$   $��Center=�6,   $	��Top=�6I   $��10q�6a    �� Padding$ 5q�6m   $��16(�6y    �� 
LineHeight$ 30��6�   
��� � � K� ZIndex$	 100��5   #   $��30��66   
������� J� Opacity$ 0Q�6U   �� 5      -   5        $��ProgressInfo��5      $��Center=�6$   $��Center=�6A   $��10q�6\   $��16(�6h   
��� � �$
��Wrap=�6�   $y�	Collapsed=�6�   � 5      -   5        $��ProgressBar��5      $��Center=�6%   $��Bottom=�6B   $	��150��6]   $y�	Collapsed=�6i   $��0 0 0 20q�6�   $��15��6�   � 5      -   5        $��MainGrid��5      $��3q�6    � RowDefinitions 5      �� 5   
   $��* �6   �� 5   
   $��* �6   �� 5   
   $��* �6   �� 5    
   $��* �6   �� 5!   
   $��* �6   �� 5"   
   $��* �6   .+�5%      �� $��0��6   ��5&   
   �� 5'      .��6   &$Файл с презентацией6'   .��5*   
   �  � ColumnDefinitions 5+      �� 5,      $��2* �6   �� 5-      $��1* �6   .+�50      �� -   5        $��PresentationFileName��50   1   $��5q�6   $��0��6!   $
��Wrap=�6M   $��15��6b   �� 51      -   5        $��0 5 5 5q�51      $��1��6$   $��Выбрать��6K   52      53   	   �� 55      $��1��6   ��56   
   �� 57      .��6   )'Название презентации6(   .��5:   
   �� -   5        $��PresentationName��5:      $��50��6*    �� AcceptsReturn .  �� TextWrapping$
 Wrap=�6K   $��5 0 5 0q�6_   5;   	   �� 5=      $��2��6   ��5>   
   �� 5?      .��6   +)Заголовок презентации6)   .��5B   
   �� -	   5        $��PresentationTitle��5B      $��50��6+    . $
 Wrap=�6L   $��5 0 5 0q�6`   5C   	   �� 5E      $��3��6   ��5F   
   �� 5G      .��6   )'Url news (Необязательное)6-   .��5J   
   �� -
   5        $��UrlNews��5J      $��50��6!    . $
 Wrap=�6B   $��5 0 5 0q�6V   5K   	   �� 5M      $��4��6   ��5N   
   �� 5O      .��6   +)Категория презентации6)   .��5R   
   �� -   5        $��CategorieComboBox��5S   #    �� SelectedValuePath$	 Key��5R   �    �� DisplayMemberPath$ Value��6�   $��5q�6�    �� SelectedIndex$ 0��5S      ��5R      ��  �� RelativeSource ��  �� AncestorType$ WindowL� �� Mode$ FindAncestor=� �� Path$ 
Categories�5S   #   5T   	   ]� 5V      $��5��6   $��5q�6   ��. .+�5X   
   �� -   5        $��30��5X      $��10 0 0 0q�6   $��30��6/   $��Right=�6:   .��5Y      �� $��/Resource/settings.png��6   5Z      �� 5[   
   -   5        $1��+Распарсить презентацию��5[      $��30��6I   $��Stretch=�6U   5\   	   5]      5^      5_      