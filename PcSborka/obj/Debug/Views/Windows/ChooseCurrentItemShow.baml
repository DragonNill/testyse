   M S B A M L   `   `   `  ����C  ?PcSborka, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null?$clr-namespace:PcSborka.Views.WindowsPcSborka.Views.Windows  5   	   2    ,PcSborka.Views.Windows.ChooseCurrentItemShow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   6local$clr-namespace:PcSborka.Views.Windows     5   	   -   5          � Title$  ChooseCurrentItemShow��5   	   $	��450��6'   $	��800��64   $	��400��6@   $	��750��6P   .��5	      �  � RowDefinitions 5
   
   �� 5      $
��auto �6   �� 5      $��1* �6   �� 5      $
��auto �6   .+�5   
   �  � ColumnDefinitions 5      �� 5      $��1* �6#   �� 5      $��2* �6#   �� 5      $��2* �6#    5      �� 5      �� 5      .+�5      �� -   5         ��Name$ BackWindow_button��5      $��10(�60   $��25��6>   $��5q�6J   $��Italic$�6V   $
��Bold"�6i   $
��Left=�6{   $	��Top=�6�   $��
Назад��6�    � Cursor$
 Hand��6�   6�   �� 5      -   5         ��Name$ Find_textBlock��5      $��Bottom=�60   $��14(�6K   $��5q�6Y   $��Italic$�6d   $
��Bold"�6w   $��1��6�   
Поиск6�   �� 5      $��Bottom=�6   $��14(�63   $��5q�6A   $��Italic$�6L   $
��Bold"�6_   $��2��6q   Сортировка6�   �� 5      -   5         ��Name$ Find_textBox��5      $��1��6,   $��1��69   $��5q�6I   $��Italic$�6U   $
��Bold"�6h   6�   �� 5      -   5         ��Name$ Sorting_comboBox��5      $��1��61   $��2��6>   $��5q�6N   $��Italic$�6Z   $
��Bold"�6m   6�   5      �� 5   
   -   5         ��Name$ CurrentItem_listView��5      $��Center=�61   $��Center=�6U   $��1��6w   6�   �� 5    
   -   5        $ AddToChoose_button��5       $��10(�6-   $��25��6;   $��5q�6G   $��Italic$�6S   $
��Bold"�6f   $
��Left=�6x   $��Добавить��6�   $��2��6�   $y�Hidden=�6�   $
 Hand��6�   �� 5!   
   -   5        $ AddNewItem_button��5!      $��10(�6,   $��25��6:   $��5q�6F   $��Italic$�6R   $
��Bold"�6e   $��Right=�6w   $<��6Добавить новые комплектующие��6�   $��2��6�   $y�Hidden=�6�   $
 Hand��6�   5"      5#      