   M S B A M L   `   `   `  ����D  @AliParser, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null%clr-namespace:AliParser	AliParser  5   	       AliParser.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   )localclr-namespace:AliParser     5   	     � Title$  	AliParser��5   	   $��489.793��6   $��694.822��6,   .��5	      � $��0,0,2,0q�6    � RowDefinitions 5
   
   �� 5      �� 5      $	��16* �6   .+�5   
   �� -   5         ��Name$ UrlText��5      $��Center=�6#   $��30��6E    �� TextWrapping$
 Wrap=�6Q   $��	Enter url��6e   $	��Top=�6v   $��14(�6�   $ �� HorizontalScrollBarVisibility$ Visible=�6�   $��0,0,74,0q�6�   $��0q�6�   $��2��6�   �� 5   
   -   5         ��Name$ StartParseBtn��5      $��Start��6(   $��Right=�68   $��0q�6T   $	��Top=�6_   $��75��6w   $��30��6�   
�������$��2��6�   &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ DataGrid��5      $��0,49,0,0q�6%   
�������$��1��6J   (  "System.Windows.Controls.WebBrowser  5   
   -   5          Name$ wb��5      $
��Left=�6!   $	��100��6<   $��0,334.154,0,-2q�6I   $	��Top=�6a   $	��100��6y   $y�Hidden=�6�    J� Opacity$ 0.025Q�6�   $��1��6�   �� 5   
   $��Средняя цена��6   $
��Left=�6(   $��10,8,0,0q�6C   $��1��6U   $	��Top=�6b   $��90��6z   �� 5   
   $!��Не рассчитано
��6   $
��Left=�62   $��	10,23,0,0q�6M   $��1��6`   $	��Top=�6m   $	��215��6�   $��26��6�   �� 5   
   $
��Left=�6   $��	105,8,0,0q�6.   $��1��6A   $	��Top=�6N   $	��120��6f   $��20��6r   	 �� SelectedIndex$	 0��6~   .@�5      �� $)��#Считать только мин.��6   �� 5      $:��4Считать мин. и макс. отдельно��6   �� 5      $C��=Считать среднее между мин. и макс.��6   5      �� 5   
   -   5        $��Экспорт в Excel��5      $
��Left=�6+   $��	576,8,0,0q�6F   $��1��6Y   $	��Top=�6f   $��99��6~   $��36��6�   6�   5      5      