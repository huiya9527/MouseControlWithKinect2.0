   M S B A M L   `   `   `  ����  KinectV2MouseControl%    KinectV2MouseControl.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   -   5          � Title$  Kinect v2 Mouse Control��5   	   $��392.197��6)   $��458.276��6:    � 
ResizeMode$ CanMinimize=�6J    � WindowStartupLocation$ CenterScreen=�6c   .��5      � .+�5   
   � -   5        $��0,0,0.5,-0.5q�5   5   .+�5      �� -   5         ��Name$ MouseSensitivity��5      $��10,40,9.5,0q�6/   $	��Top=�6D   $��26��6\    � Maximum$ 6Q�6h    � Minimum$ 1Q�6�    � Value$	 3.5Q�6�   �� 5      $��Mouse Sensitivity��6   $��10,10,117.5,0q�60   $	��Top=�6G   �� 5	      -   5        $ PauseToClickTime��5	      $��10,111,9.5,0q�6/   $	��Top=�6E   $��26��6]   $ 5Q�6i    � SmallChange$
 0.01Q�6u   $ 2Q�6�   $	 0.3Q�6�   �� 5
      $"��Pause-To-Click Time Required��6   $��10,81,117.5,0q�6;   $	��Top=�6R   �� 5      -   5         ��Name$ txtTimeRequired��5      $��23��6/   $��0,81,20.5,0q�6;   	 �� TextWrapping$
	 Wrap=�6P   $	��Top=�6d   $��2��6|   $��15(�6�   
 �� TextAlignment$
 Right=�6�   $��Right=�6�   $��77��6�   6�   �� 5      -   5        $ txtMouseSensitivity��5      $��23��63   $��0,12,20.5,0q�6?   $
	 Wrap=�6T   $	��Top=�6h   $	��3.5��6�   $��15(�6�   $
 Right=�6�   $��Right=�6�   $��77��6�   6�   �� 5      -   5        $ PauseThresold��5      $��10,182,9.5,0q�6,   $	��Top=�6B   $��26��6Z   $	 160Q�6f   $
 0.01Q�6t   $ 60Q�6�   $ 10Q�6�   �� 5      $��Pause Movement Thresold��6   $��10,152,117.5,0q�66   $	��Top=�6N   �� 5      -   5        $ txtPauseThresold��5      $��23��60   $��0,152,20.5,0q�6<   $
	 Wrap=�6R   $	��Top=�6f   $��60��6~   $��15(�6�   $
 Right=�6�   $��Right=�6�   $��77��6�   6�   �� 5      -	   5         ��Name$ 
btnDefault��5      $��Default��6)   $��Right=�6;   $��0,323,15.5,0q�6W   $��74��6m   $��24��6�   $	��Top=�6�   �� 5      -
   5         ��Name$ 
chkNoClick��5      $!��No clicks, move cursor only��6+   $
��Left=�6Q   $��260,295,0,0q�6l   $��16��6�   $	��Top=�6�   6�   	� 5      -   5         	�Name$ rdiGrip��5      $��Grip Gesture��6+   $
��Left=�6B   $��
18,297,0,0q�6]   $	��Top=�6q    	� 	GroupName$ GestureMode��6�   6�   	� 5      -   5        $ rdiPause��5      $��Pause To Click��6,   $
��Left=�6E   $��134,297,0,0q�6`   $	��Top=�6u   $ GestureMode��6�   6�   �� 5      -   5        $ CursorSmoothing��5      $��10,253,9.5,0q�6.   $	��Top=�6D   $��26��6\   $
 0.95Q�6h   $
 0.01Q�6w   $	 0.2Q�6�   6�   �� 5      $��Cursor Smoothing��6   $��10,223,117.5,0q�6/   $	��Top=�6G   �� 5      -   5        $ txtCursorSmoothing��5      $��23��62   $��0,223,20.5,0q�6>   $
	 Wrap=�6T   $	��Top=�6h   $	��0.2��6�   $��15(�6�   $
 Right=�6�   $��Right=�6�   $��77��6�   6�   �� 5      $8��2Developed by Jingzhou Chen  (http://TangoChen.com)��6   $
��Left=�6Q   $��
10,323,0,0q�6l   $	��Top=�6�   5      5      5      