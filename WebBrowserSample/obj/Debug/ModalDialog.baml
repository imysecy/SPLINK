   M S B A M L   `   `   `  ����  WebBrowserControlSample)    #WebBrowserControlSample.ModalDialog   5      \ XPresentationFramework, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      8x,http://schemas.microsoft.com/winfx/2006/xaml    5      Hmc;http://schemas.openxmlformats.org/markup-compatibility/2006    5      >d2http://schemas.microsoft.com/expression/blend/2008    5      .��5      �   � DataContext  6   ��  �� ElementName$
 root��.+�5	   
   �� ���	#90000000y�6(   �� *
Visibility+._�5
      �� 
���   �$��1q�6)   
������� �� CornerRadius$ 	10,0,10,0��5      $��Center=�6*   $��Center=�5       J� BitmapEffect 5      F� 5       F� Color$ Black��6-    F� Opacity$	 0.5Q�5       F� 	Direction$	 270Q�6!    F� ShadowDepth$	 0.7Q�5      ._�5      � $��10q�6    � RowDefinitions 5      �� 5      �� 5      $
��Auto �6(   .+�5      �� -   5        	 ��Name$	 MessageTextBlock��5       $
��Wrap=�5      $��5q�6)   ��5      �� *	Message+5   )   A� 5      $��1��6"   $��5q�6/   
 A� Columns$
 2��5      $��Center=�6'   $��Bottom=�5      .+�5      �� -   5         ��Name$ OkButton��5   !   $��Ok��5      $��2q�6(   �� 5      -   5        $ CancelButton��5   !   $��Cancel��5!      $��2q�6,   5"      5#      5$      5%      5&      5'      