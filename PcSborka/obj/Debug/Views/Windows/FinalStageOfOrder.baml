   M S B A M L   `   `   `  ����C  ?PcSborka, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null?$clr-namespace:PcSborka.Views.WindowsPcSborka.Views.Windows  5   	   .    (PcSborka.Views.Windows.FinalStageOfOrder   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   6local$clr-namespace:PcSborka.Views.Windows     5   	     � Title$  FinalStageOfOrder��5   	   $	��300��6#   $	��550��60   $	��275��6<   $	��500��6L   .��5	      �  � RowDefinitions 5
   
   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6   .+�5   
   �  5      �� 5      $
��auto �6    �� 5      $��1* �6    .+�5      ��  �� TextAlignment$ Center=�6   $��Center=�6/   $	��Top=�6L   $��16(�6d   $��5q�6r   $��Italic$�6}   $
��Bold"�6�   ;9Добро пожаловать в меню заказа,�� 6�   ?=теперь осталось только заполнить�� 6�   оставшиеся поля6
  �� 5      -   5         ��Name$ Back_button��5      $��12(�6+   $��5q�69   $��Italic$�6E   $
��Bold"�6X   $
��Left=�6j   $	��Top=�6�   $��
Назад��6�   6�   � 5      $��1��6    � ColumnDefinitions 5      �� 5      $��1* �6'   �� 5      $��4* �6'   .+�5      �� $��0��6   $��Right=�6,   $��Center=�6H   $��16(�6c   $��5q�6q   $��Italic$�6|   $
��Bold"�6�   Почта:6�   �� 5      -   5         ��Name$ Email_textBox��5      $��16(�61   $��5q�6?   $��Italic$�6J   $
��Bold"�6]   $��1��6o   5      5       � 5!   
   $��1��6    5"      �� 5#      $��1* �6#   �� 5$      $��4* �6#   .+�5&      �� $��0��6   $��Right=�6(   $��Center=�6D   $��16(�6_   $��5q�6m   $��Italic$�6x   $
��Bold"�6�   Адрес:6�   �� 5'      -   5        $ Address_textBox��5'      $��16(�6/   $��5q�6=   $��Italic$�6H   $
��Bold"�6[   $��1��6m   5(      � 5)   
   $��2��6    5*      �� 5+      $��1* �6#   �� 5,      $��4* �6#   .+�5.      �� $��0��6   $��Right=�6(   $��Center=�6D   $��16(�6_   $��5q�6m   $��Italic$�6x   $
��Bold"�6�   Телефон:6�   �� 5/      -   5        $ Phone_textBox��5/      $��16(�6-   $��5q�6;   $��Italic$�6F   $
��Bold"�6Y   $��1��6k    �� 	MaxLength$ 11��6�    � Cursor$ Arrow��6�   50      �� 51   
   -   5        $ CreateOrder_button��51      $��3��6-   $��12(�6:   $��5q�6H   $��Italic$�6T   $
��Bold"�6g   $��Center=�6y   $��Center=�6�   .��6�   Сделать заказ6�   52      53      