   M S B A M L   `   `   `  ����Z  VWPFToolkit.Extended, Version=1.8.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f4R<clr-namespace:Xceed.Wpf.Toolkit;assembly=WPFToolkit.ExtendedXceed.Wpf.Toolkit  5       MPSI   MPSI.MainWindow   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Mxctk<clr-namespace:Xceed.Wpf.Toolkit;assembly=WPFToolkit.Extended     5      -   5          �Name$  windows1��5   O    � Title$ 
MainWindow��5      $
��1024��6   $
��1280��6&    � 
ResizeMode$ CanResizeWithGrip=�63    � WindowStartupLocation$ CenterScreen=�6R   .��5      �� -   5         ��Name$ Canvas1��5      $��Stretch=�6   $��Stretch=�6=   $��0,0,-8,5q�6Y   .+�5   
   �� -   5         ��Name$ Bouton1��5      $��Lancer l'animation��6"   $��32��6?    �� Left$ 25��6K   
 �� Top$ 27��6\   $	��112��6l   �� 5   
   -   5         ��Name$ Textbox1��5      $��23��6$   	 �� TextWrapping$
	 Wrap=�60   $��TextBox��6D   $	��480��6S   $ 25��6_   $ 4��6p   �� 5	   
   -   5        
 ��Name$
 Scroll1��5	      $��8��6%   $	��Top=�60   $��Center=�6H    �� Orientation$ 
Horizontal=�6e   $
��1270��6~    � SmallChange$ 1Q�6�   $��0q�6�   �� 5
   
   -   5        $ Panneau��5
      $	��188��6"   $ 25��6/   $	 764��6@   $	��814��6Q   .+�5      �� 
�������$	��188��6%   
���   �$	��814��6A    � BindingGroup 5      &   System.Windows.Data.BindingGroup  5      5      �� 5      -   5        $ Textbox2��5      $��22��6(   $
	 Wrap=�64   $��1��6H   $��32��6Q   $ 16��6\   $ 74��6l   �� 5      -   5         ��Name$ Cbbx1��5      $	��546��6&   $ 52��62   $ 70��6B   �� 5      $��t:��6   $��24��6!   $��18��6-   $��Center=�68   $��Center=�6U   {�.  $ 14��6�   $ 48��6�   �� 5      $��en secondes��6   $��24��6*   $	��114��66   $��Center=�6B   $��Center=�6_   {�.  $ 14��6�   $	 106��6�   �� 5      $��
Période :��6   $��24��6(   $��66��64   $��Center=�6?   $��Center=�6\   {�.  $ 50��6�   $ 14��6�   �� 5      $"��Seuil minimal de cohérence:��6   $��24��6:   $ 14��6F   $ 92��6W   $	��168��6g   $��Center=�6s   $��Center=�6�   {�.  6�   %   Xceed.Wpf.Toolkit.MaskedTextBox  5      -	   5          Name$ MTXTBox2��5   !   $	 176��63   $
	 Wrap=�6E   $
��0,20��6Y   $ 94��6e   $��34��6u      Mask$
 0.00��6�   )   #Xceed.Wpf.Toolkit.AutoSelectTextBox   AutoSelectBehavior$ OnFocus=�6�   5      5      5      