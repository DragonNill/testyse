   M S B A M L   `   `   `  ����C  ?PcSborka, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null?$clr-namespace:PcSborka.Views.WindowsPcSborka.Views.Windows  5   	   0    *PcSborka.Views.Windows.FinishToOrderWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   6local$clr-namespace:PcSborka.Views.Windows     5   	   -   5          � Title$)  #Завершенние заказа��5   	   $	��450��6$   $	��800��61   .��5	      �  � RowDefinitions 5
   
   �� 5      $
��1.5* �6   �� 5      $��1* �6   �� 5      $��1* �6   .+�5   
   �  5      �� 5      �� 5      �� 5      �� 5      .+�5      �� -   5         ��Name$ Back_button��5      $��12(�6+   $��5q�69   $��Italic$�6E   $
��Bold"�6X   $
��Left=�6j   $	��Top=�6�   $��
Назад��6�   6�   �� 5      $��0��6    �� TextAlignment$ Center=�6%   $��Center=�6<   $	��Top=�6Y   $��16(�6q   $��5q�6   $��Italic$�6�   $
��Bold"�6�   &$Информация о заказе6�   �� 5      -   5         ��Name$ Address_textBlock��5      $��1��63   $ Center=�6@   $��Center=�6W   $	��Top=�6t   $��14(�6�   $��5q�6�   $��Italic$�6�   $
��Bold"�6�   6�   �� 5      -   5        $ Email_textBlock��5      $��2��61   $ Center=�6>   $��Center=�6U   $	��Top=�6r   $��14(�6�   $��5q�6�   $��Italic$�6�   $
��Bold"�6�   6�   �� 5      -   5        $ Phone_textBlock��5      $��3��61   $ Center=�6>   $��Center=�6U   $	��Top=�6r   $��14(�6�   $��5q�6�   $��Italic$�6�   $
��Bold"�6�   6�   5      � 5   
   $��1��6    � ColumnDefinitions 5      �� 5      �� 5      .+�5!      �� -   5        $ FinishTheComputer_button��5!      $��13(�67   $��5q�6E   $��Italic$�6Q   $
��Bold"�6d   $��Center=�6v   $��Center=�6�   $+��%Дособрать компьютер��6�   6�   �� 5"      -   5        $ FinishThePeriphery_button��5"      $��13(�68   $��5q�6F   $��Italic$�6R   $��1��6e   $
��Bold"�6u   $��Center=�6�   $��Center=�6�   {�.  $-��'Посмотреть периферию��6�   6�   5#      � 5$   
   $��2��6    5%      �� 5&      �� 5'      �� 5(      .+�5*      �� $��0��6   $ Center=�6%   $��Center=�6<   $	��Top=�6Y   $��14(�6q   $��5q�6   $��Italic$�6�   $
��Bold"�6�   ��Сумма за работу, над сборкой, выбирает сотрудник, она добавится к итоговой сумме сборке6  �� 5+      $��1��6    �� Orientation$ 
Horizontal=�6&   $��Center=�6?   $��Center=�6Z   .+�5,      �� -   5        $ AllInCost_textBlock��5,      $��14(�6:   $��5q�6H   $��Italic$�6S   $
��Bold"�6f   Цена сборки ? +6�   �� 5-      -	   5         ��Name$ FinalCost_textBox��5-   )    �� 	MaxLength$ 5��6   $	��100��6D   $��12(�6P   $��5q�6^   $��Italic$�6i   $
��Bold"�6|   	 � Cursor$	 Arrow��6"  �� 5.      -
   5        $% FinalComputerAssembly_textBlock��5.      $ Center=�6E   $��14(�6]   $��5q�6k   $��Italic$�6v   $
��Bold"�6�   " = Финальная сумма?6�   5/      �� 50      -   5        $ EndToOrder_button��50      $��2��60   $
��Left=�6=   $��Center=�6Y   $��12(�6t   $��5q�6�   $��Italic$�6�   $
��Bold"�6�   $y�Hidden=�6�   $
	 Hand��6�   .��6�   Завершить заказ6  51      52      53      