   M S B A M L   `   `   `  ����G  CWPF_SYSTools, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:WPF_SYSToolsWPF_SYSTools  5   	       WPF_SYSTools.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   b ^MaterialDesignThemes.Wpf, Version=4.0.0.2422, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   i eHardcodet.Wpf.TaskbarNotification, Version=1.0.18.0, Culture=neutral, PublicKeyToken=b9e9fbe9b6908adc)tb http://www.hardcodet.net/taskbar  5   	   &localclr-namespace:WPF_SYSTools  5   	   -   5          � Title$  SYSTools��5
   	   $	��650��6   $
��1100��6'    � WindowStyle$
 None=�64    � AllowsTransparency . ��.  |�.  $y�Visible=�6�   ��6a   Z�  � 	Resources 5         ScrollBarThumb �� 
TargetType �� Property �� Value �� 
TargetType �� CornerRadius ! HorizontalScrollBarPageButton  VerticalScrollBarPageButton   VerticalScrollBarPageButton2  for_scrollbar	 �Name
 J� SnapsToDevicePixels � RowDefinitions �� PageUpCommand ]�Name J� IsMouseOver ]� DecreaseRepeatButton ]� IncreaseRepeatButton �� PageDownCommand ]� Thumb �� Triggers P� Property P� Value �� PageLeftCommand �� PageRightCommand  for_scrollviewer � Cursor ��Name �� ScrollableHeight � Maximum �� Orientation �� ViewportHeight �� ViewportSize �� VerticalOffset  � Value*! �� #ComputedVerticalScrollBarVisibility" �� ScrollableWidth# �� HorizontalOffset$ �� ViewportWidth,% �� %ComputedHorizontalScrollBarVisibility& �� Triggers' :� RoutedEvent( �� 
TargetName) �� TargetProperty	* W� To+ g� Duration, g� 	BeginTime- :� 
SourceName. �� BasedOn��%  &	        &	     &	 �    &	 �    &	 x    7  7  7  7   7  7  7  7   &	 K    7  7  '�� X    7  '�� z    7  �� 5   
   " n�6'   .�5      ��  �@�OverridesDefaultStyle . 66   �� 5       � �� .  6*   �� 5       � �� 5      �� 5      " n�6&   .R�5      � .+�5      �� ���	#FF646464$ 4��6<   5      5      5      5      �� 5   
   " ��66   .�5      ��  �@�OverridesDefaultStyle . 66   �� 5       � ��
 ���� 6+   �� 5       � |� .  6*   �� 5        � �� .  6*   �� 5!       �@J�Opacity$	 0.2Q�6(   �� 5"       � �� 5#      �� 5$      " ��6&   .R�5%      �� #��z��6$   #��z��6H   #��z��6h   5&      5(      5)      �� 5*   
   " ��64   .�5+      ��  �@�OverridesDefaultStyle . 66   �� 5,       � ��
 ���� 6+   �� 5-       � |� .  6*   �� 5.       � �� .  6*   �� 5/       �@J�Opacity$	 0.2Q�6(   �� 50       � �� 51      �� 52      " ��6&   .R�53      �� #��z��6$   #��z��6H   #��z��6h   54      56      57      �� 5:   
   " ��65   .�5;      ��  �@�OverridesDefaultStyle . 67   �� 5<       � ��
 ���� 6+   �� 5=       � |� .  6*   �� 5>       � �� .  6*   �� 5?       �@J�Opacity$ 0Q�6)   �� 5@       � �� 5A      �� 5B      " ��6&   .R�5C      �� ���	#90000000$��0��65   $��0��6?   5D      5F      5G      �� 5H   
   " ��6&   .�5I      ��  �@��IsPressAndHoldEnabled .  6=   �� 5J       �@��IsFlicksEnabled .  67   �� 5K       � ��
 ���� 6+   �� 5L       � ��$ 0,1,1,6q�6'   �� 5M       � ��$ 10��6'   �� 5N       � ��$ 5��6*   �� 5O       �@J�Opacity$	 0.2Q�6*   �� 5P       � �� 5Q      �� 5R      " ��6&   .R�5S      � $	 Bg��6   
 .  5T      �� 5U   "   $
��auto �60   6@   �� 5V   "   $��* �60   6=   �� 5W   "   $
��auto �60   6@   .+�5Y      �� $��0��6,   8��  #��Z 6o   ]� 5Z      $ 
PART_Track��6$   $��1��6:   #{�z 6G   ��.  5[   "   �� 5\   &   8�� #��Z 6h    5^   "   �� 5_   &   8�� #��Z 6h    5a   "   n� 5b   &   8�� 6,   5d      �� 5e      $��2��6+   8�� #��Z 6n   5f      5g      5i       5j      P� 5k       �@��Orientation$ 
Horizontal=�61   . �5l      ��  � ��
 ���� 64   �� 5m       � ��$ 1,0,6,1q�6/   �� 5n       � ��$ 5��61   �� 5o       � ��$ 5��63   �� 5p       � ��$
 Auto��61   �� 5q       �@J�Opacity$	 0.2Q�60   �� 5r       � �� 5s      �� 5t      " ��6.   .R�5u   "   � $	 Bg��6'   
 . .+�5v   &   ]� $ 
PART_Track��6,   #{�z 6@    5w   *   �� 5x   .   8�� #��Z 6s    5z   *   �� 5{   .   8�� #��Z 6s    5}   *   n� 5~   .   8�� 64   5�   '   5�   #   5�      5�      5�      5�      �� 5�   
   " ��5�      .�5�      ��  � ��
 �����6,   �� 5�       � ��$ 0q�60   �� 5�       � �� 5�      �� 5�      " ��6&   .R�5�      �� #��z��6!   #��z��6M   
 . ._�5�      � #��z��6#   .+�5�   "   �� # z 6:   #��z��6\   #��z��6   �� 5�   "   $ PART_VerticalScrollBar��6,   $��Right=�5�   (   # z 5�   (   $ Vertical=�5�   (   8��  # z 5�   (   #  z 5�   (   #y�z! 5�   (   �� 5�   "   $ PART_HorizontalScrollBar��6,   # z" 5�   (   $ 
Horizontal=�5�   (   8�� $��Bottom=�5�   (   #  z# 5�   (   # z$ 5�   (   #y�z% 5�   (   5�      5�      & 5�      :� 5�      $' ScrollChanged��6+   .[�5�   "   �� .��5�   &   �� .��5�   *   W� $( PART_VerticalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.8X�5�   )   $+ 0:0:1D�5�   )   W� 5�   *   $( PART_VerticalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.2X�5�   )   $+ 0:0:1D�5�   )   $, 0:0:1X�5�   )   W� 5�   *   $( PART_HorizontalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.8X�5�   )   $+ 0:0:1D�5�   )   W� 5�   *   $( PART_HorizontalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.2X�5�   )   $+ 0:0:1D�5�   )   $, 0:0:1X�5�   )   5�   '   5�   #   5�      :� 5�      $' 
MouseEnter��6+   $- PART_VerticalScrollBar��5�   '   .[�5�   "   �� .��5�   &   �� .��5�   *   W� $( PART_VerticalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.8X�5�   )   $+ 0:0:0.7D�5�   )   5�   '   5�   #   5�      :� 5�      $' 
MouseLeave��6+   $- PART_VerticalScrollBar��5�   '   .[�5�   "   �� .��5�   &   �� .��5�   *   W� $( PART_VerticalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.2X�5�   )   $+ 0:0:0.7D�5�   )   5�   '   5�   #   5�      :� 5�      $' 
MouseEnter��6+   $- PART_HorizontalScrollBar��5�   '   .[�5�   "   �� .��5�   &   �� .��5�   *   W� $( PART_HorizontalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.8X�5�   )   $+ 0:0:0.7D�5�   )   5�   '   5�   #   5�      :� 5�      $' 
MouseLeave��6+   $- PART_HorizontalScrollBar��5�   '   .[�5�   "   �� .��5�   &   �� .��5�   *   W� $( PART_HorizontalScrollBar��5�   )   $) Opacity�5�   )   $	* 0.2X�5�   )   $+ 0:0:0.7D�5�   )   5�   '   5�   #   5�      5�      5�      5�      �� 5�   
   " ��6   8.   6'   �� 5�   
   " ��6   8.   6*   5�      .��5�      � $��0q�6   .+�5  
   3  -Hardcodet.Wpf.TaskbarNotification.TaskbarIcon  -   5        /   
IconSource$/ Resources/SYSTools.ico��5     0 � ContextMenu0 5     �� 5     
�������.@�5     m� -   5        1 m�Name$1 OpenSYS��5     $��打开SYSTools��60   m� 5	     -   5        $1 ExitSYS��5	     $��退出应用��60   5
     5     �� 5  
   -   5        2 ��Name$2 SP_Sub��5     $
��Left=�6%   $	��650��6@   $	��Top=�6M   $	��100��6e   .+�5     �� -   5        ���#E0E0E0$
��Left=�5  +   $	��650��6F   $	��Top=�6S   $��88��6k   $ 	0,20,20,0��6v   3 J� Effect3 5     3  -System.Windows.Media.Effects.DropShadowEffect  5     4   
BlurRadius$4 10Q�6'   5   ShadowDepth$5 0Q�67   5     5     �� 5     -   5        6 ��Name$6 HomeLOGO��5     7 � ToolTip$7 工具箱LOGO 欢迎页��6&   $	��Top=�6<   8 J� RenderTransformOrigin$8 0.5,0.5&�6T   $!��Resources/工具箱Logo.png��6t   $
��Left=�6�   $��40��6�   $��50��6�   $��19,-625,0,0q�6�   �� 5     ���	#FAFAFAFA$��1q�6-   $
��Left=�6A   $��2��6\   $	��Top=�6g   $��68��6   $ 0.5,0.5,0.5,0.5��6�   $��10,-570,0,0q�6�   �� 5     -   5        9 ��Name$
9 Info��5      & "MaterialDesignFloatingActionButton#��[ 6#   $7 电脑硬件信息识别��6_   $
��Left=�5     $	��Top=�6,   $��68��6D   $��68��6P   ���	#FFEEEEEE���	#FF616161��.  $8 0.469,0.376&�6�   $��10,-550,0,0q�6�   |�.  .��5     '  !MaterialDesignThemes.Wpf.PackIcon  :   Kind$: 
DesktopMac=�6*   $��30��6<   $��30��6H   5     �� 5!     -	   5        $
9 Apps��5!     #��[ 6#   $7 常规电脑硬件工具��6_   $
��Left=�5"     $	��Top=�60   $��68��6H   $��68��6T   ���	#FFEEEEEE���	#FF616161��.  $��10,-462,0,0q�6�   |�.  .��5#       $
: Apps=�6*   $��30��66   $��30��6B   5$     �� 5&     -
   5        $
9 Over��5&     #��[ 6#   $7 超频或BIOS修改工具��6_   $
��Left=�5'     $	��Top=�6,   $��68��6D   $��68��6O   ���	#FFEEEEEE���	#FF616161��.  $��10,-372,0,0q�6�   |�.  .��5(       $	: Owl=�6*   $��30��65   $��30��6A   5)     �� 5+     -   5        $9 	Benchmark��5+     #��[ 6(   $7 压力测试工具��6d   $
��Left=�5,     $	��Top=�6,   $��68��6D   $��68��6P   ���	#FFEEEEEE���	#FF616161��.  $��10,-282,0,0q�6�   |�.  .��5-       $: Whatshot=�6*   $��30��6:   $��30��6F   5.     �� 50     -   5        $9 	WActivate��50     #��[ 6(   $7 Windows Activate��6d   $
��Left=�51     $	��Top=�6,   $��68��6D   $��68��6P   ���	#FFEEEEEE���	#FF616161��.  $��
10,-80,0,0q�6�   |�.  .��52       $	: Cat=�6*   $��30��65   $��30��6A   53     54     �� 56  
   -   5        $2 SP_CMbutton��56     $��Right=�6*   $��26��6F   $��0q�6R   $	��Top=�6]   $��90��6u   .+�57     �� -   5        $9 Minus_Button��57     $7 	最小化��6U   $
��Left=�58     $��45��6,   $��26��6:   $ Arrow��6F   ���	#FFA0A0A0���	#FFA0A0A0��.  $��Center=�6�   $��0q�6�   |�.  ��57  +   0  *MaterialDesignThemes.Wpf.PackIconExtension  *Minus+58  �   �� 59     -   5        $9 Close_Button��59     $7 关闭��6U   $��45��6c   $��26��6o   $ Arrow��6{   ���	#FFDD2C00���	#FFDD2C00��.  $	��Top=�6�   $��Right=�6�   $��	0,-26,0,0q�6�   |�.  ��6+     *Close+6  5:     �� 5<  
   -   5        $2 SP_Main��5<     $
��Left=�6&   $
��1000��6A   $��	100,0,0,0q�6N   .+�5=     �� ; �� Orientation$; 
Horizontal=�6   $��62.934��62   $��Center=�6B   $��291.918,68,291.918,0q�6]   $��Center=�6{   .+�5>     ?  9MaterialDesignThemes.Wpf.Transitions.TransitioningContent  C  =MaterialDesignThemes.Wpf.Transitions.TransitioningContentBase<   OpeningEffect< 66   D  >MaterialDesignThemes.Wpf.Transitions.TransitionEffectExtension  *FadeIn+.��5?     �� 
���   �$��Center=�63   $	��0,4q�6P   $��Center=�6]   �� 6y   .F�5@  &     -   5        =  Name$= UserSEgg��5@  >   $��50��6P   $	: Cat=�6\   $��50��6g   $��Center=�6r   $��	0,0,0,-10q�6�   $��Center=�6�   $��Center=�6�   $��Center=�6�   $��0q�6  $7  ��6  |�.  6  5A  #   �� 66   $� ��6:   �� 6E   .F�5B  &   �� $��Center=�60   $#��欢迎使用SYSTools工具箱��6K   
���   �$	��32 (�6u   $��Center=�6�   > �� Padding$> 0,0,0,8q�6�   5C  #   67   5D     5G     #  MaterialDesignThemes.Wpf.Card  5H     -   5        ?  Name$? 
Card_About��5H  "   #��� ��66   ���	#DDFFFFFF$��8q�6�   @   UniformCornerRadius$@ 6Q�6�   $��0,45,0,0q�6�   $��Center=�6�   $	��Top=�6�   $	��942��6  $	��299��6  .��5I       $��2,-1.5,1.5,1.5q�6o   < 66     *FadeIn+.��5J     � $��10q�6   A � ColumnDefinitionsA 5K     �� 5L     $
��231* �6/   �� 5M     $
��231* �6/   �� 5N     $
��231* �6/   �� 5O     $
��231* �6/   .+�5Q       $: MicrosoftWindows=�62   $��Center=�6K   $	��Top=�6h   $	��100��6�   $	��100��6�   $��65.75,28.3,65.75,0q�6�   |�.  6�   �� 5R     B �� TextAlignment$B Center=�6$   $��Center=�6A   $��Center=�6\   $��15(�6y   $��211.5��6�   $��51.884��6�   $��10,167.458,10,30.658q�6�   �� 6�   $�简单易用��6�   �� 6�   $� 专为Windows设计��6�   �� 6�   �� 6�   �� 6  $�基于��6  �� 6  $�.Net 4.6制作��6  �� 6*  67    5S     $: MaterialDesign=�62   $��Center=�6I   $	��Top=�6f   $	��100��6~   $	��100��6�   $��1��6�   $��65.75,28.3,65.75,0q�6�   |�.  6�   �� 5T     $B Center=�6$   $��Center=�6A   $��Center=�6\   $��15(�6y   $��211.5��6�   $��51.884��6�   $��10,167.458,10,30.658q�6�   $��1��6�   �� 6�   $	�WPF��6�   �� 6�   $	� UI��6�   �� 6�   �� 6  $�Material Design��6  �� 6  $� In XAML��6#  66    5U     $ : AccountQuestionMarkOutline=�62   $��Center=�6U   $	��Top=�6r   $	��100��6�   $	��100��6�   $��2��6�   $��65.75,28.3,65.75,0q�6�   |�.  6�   �� 5V     $B Center=�6$   $��Center=�6A   $��Center=�6\   $��15(�6y   $��211.5��6�   $��51.884��6�   $��10,167.458,10,30.658q�6�   $��2��6�   �� 6�   $�有疑问 请加群��6�   �� 6�   �� 6�   $�反馈问题&提出意见��6�   6    5W     $: Github=�62   $��Center=�6A   $	��Top=�6^   $	��100��6v   $	��100��6�   $��3��6�   $��65.75,28.3,65.75,0q�6�   |�.  6�   �� 5X     $B Center=�6$   $��Center=�6A   $��Center=�6\   $��15(�6y   $��211.5��6�   $��51.884��6�   $��10,167.458,10,30.658q�6�   $��3��6�   �� 6�   $�软件开源��6�   �� 6�   �� 6�   $�(��6�   �� 6  $�Don't be evil��6  �� 6  $�)��6   6,  5Y     5Z     5[     � 5\     $��Center=�6   $��Center=�60   $��28.875,0,461.125,0q�6K   .+�5^     �� -   5        $9 Webhome��5^     #��[ 6*   $'7 !软件官网 (目前尚未完成)��6f   $
��Left=�5_     $��Center=�60   $��70��6K   $��2,43.5,0,55.5q�6V   $��70��6m   ���	#FFEEEEEE���	#FF616161��.  |�.  .��5`       $: HomeLock=�6.   $��30��6>   $��30��6J   5a     �� 5c     -   5        $9 QQ��5c     #��[ 6%   $7 软件QQ群��6a   $
��Left=�5d     $��Center=�60   $��70��6K   $��162,43.5,0,55.5q�6V   $��70��6o   ���	#FFEEEEEE���	#FF616161��.  |�.  .��5e       $: Qqchat=�6.   $��30��6<   $��30��6H   5f     �� 5g     -   5        $9 Github��5g     #��[ 6)   $7 软件开源��6e   $
��Left=�5h     $��Center=�60   $��70��6K   $��252,43.5,0,55.5q�6V   $��70��6o   ���	#FFEEEEEE���	#FF616161��.  |�.  .��5i       $: Github=�6.   $��30��6<   $��30��6H   5j     �� 5k     -   5        $9 BiliBIli��5k     #��[ 6+   $7 BiliBili��6g   $
��Left=�5l     $��Center=�60   $��70��6K   $��342,43.5,0,55.5q�6V   $��70��6o   ���	#FFEEEEEE���	#FF616161��.  |�.  .��5m     �� $��Resources/BiliBili.png��6   5n     �� 5o     -   5        $9 CoolApk��5o     #��[ 6*   $7 酷安账号��6f   $
��Left=�5p     $��Center=�60   $��70��6K   $��432,43.5,0,55.5q�6V   $��70��6o   ���	#FFEEEEEE���	#FF616161��.  |�.  .��5q     �� $��Resources/Coolapk.png��6   5r     '	  !MaterialDesignThemes.Wpf.Snackbar	  5s     -   5        C 	 Name$C HomeSnackbar��5s  *   D 	  IsActiveD .  $��0,0,0,-6q�6Q   $��0q�6c   ��.  |�.  E 	  Message.E 5t     .
  (MaterialDesignThemes.Wpf.SnackbarMessage
  -   5        F 
 Name$F HomeSnackbarMessage��5t  5   G 
  ActionContent$G 好的��6R   ��.  |�.  6w   5u     5v       5w     #��� ��6"   
�������$��8q�6�   $@ 6Q�6�   $��0,-165,29,0q�6�   $��Right=�6�   $��Center=�6�   $	��400��6�   $	��150��6   $8 0.5,0.5&�6  .��5x     �  5y     �� 5z     $	��25* �6(   �� 5{     $
��109* �6(   .+�5}     �� $B Center=�6    $��Center=�6=   $��Center=�6X   $��14(�6u   $
��Wrap=�6�   $!��一言卡片(Hitokoto Card)��6�   $��0��6�   �� 5~     -   5        H ��Name$H Hitokoto��5~      $B Center=�62   $��Center=�6O   $��Center=�6j   $��15(�6�   $
��Wrap=�6�   $��	TextBlock��6�   $��1��6�   5     5�     5�     �� 5�  
   -   5        $2 SP_Info��5�     $
��Left=�6&   $
��1000��6A   $��	100,0,0,0q�6N   $y�Hidden=�6a   .+�5�     �� $
��Wrap=�6   $!��硬件与系统信息检测��6,   $
��Left=�6=   $	��230��6X   $��34��6d   
���   �$��25(�6�   $��	44,23,0,0q�6�   $
��Bold"�6�   �� 5�     -   5        $
��Left=�5�     $	��550��61   $	��Top=�6>   $	��950��6V   $ 20��6b   $��
25,20,25,0q�6�   ��5�     �� 5�     I �� Color$I #E0E0E0��6&   3 5�       5�     $4 10Q�6'   $5 0Q�67   5�     ._�5�     �� $��10q�6   .+�5�     �� -   5        J ��Name$J Info_ListBox��5�     $��15(�64   $
��Auto=�6B   $	��530��6t   |�.  5�     5�     5�     �� 5�     -   5        $9 	Info_Test��5�     #��[ 6(   $
7 ん?��6d   $��Right=�5�     $��Center=�61   $��70��6L   $��0,-110,55,0q�6W   $��70��6l   ���	#FFEEEEEE���	#FF616161��.  |�.  .��5�       $: GoogleDownasaur=�6*   $��30��6A   $��30��6M   5�     5�     �� 5�  
   -   5        $2 SP_Apps��5�     $
��Left=�6&   $
��1000��6A   $��	100,0,0,0q�6N   $y�Hidden=�6a   .+�5�     �� $
��Wrap=�6   $$��常规电脑硬件检测工具��6,   $
��Left=�6>   $	��250��6Y   $��34��6e   
���   �$��25(�6�   $��	44,23,0,0q�6�   $
��Bold"�6�   �� 5�     -    5        K ��Name$K App_SV��5�     $��0,23,0,0q�6+   |�.  $	��570��6O   .��5�     � -!   5        $	 App_Grid��5�     $	��750��6)   $��0q�66    5�     �� 5�     $	��310 �6(   �� 5�     $��10 �6(   �� 5�     $	��210 �6(   �� 5�     $��10 �6(   �� 5�     $	��210 �6(   .+�5�     �� $��0q�6   $��Center=�61   $��6��6N   $��Center=�6Y   $	��950��6t   $
��25,2q�6�   $��1��6�   ���#616161$ 2��6�   �� 5�     $��0q�6   $��Center=�61   $��6��6N   $��Center=�6Y   $	��950��6t   $
��25,2q�6�   $��3��6�   ���#616161$ 2��6�     5�     -"   5        $? HDCard��5�  *   $
��Left=�6:   $��Center=�6U   $	��150��6p   $	��150��6}   ��.  $27 ,硬盘工具(该列为正常检测和分区)��6�   $��30,60,0,100q�6�   .��5�     � $��Center=�6   $��0q�6<   $��Center=�6G   $	��130��6b   $	��130��6n   .+�5�       $: Harddisk=�66   $	��100��6F   $	��100��6S   ���#212121$��Center=�6t   $��Center=�6�   5�     5�     �� 5�     -#   5        $9 AS_SSD��5�     $
��Left=�6-   $��	210,5,0,0q�6H   $	��Top=�6[   $	��100��6s   $	��100��6   ��.  |�.  $��0q�6�   L � IsSharedSizeScopeL . $7 AS SSD Benchmark��6�   ���#E0E0E0���#E0E0E0��6!  Z� .��5�     �� $��Resources/AS SSD.png��6    $��80��6>   $��Center=�6I   $��Center=�6f   $��80��6�   5�     �� 5�     -$   5        $9 CrystalDiskInfo��5�     $
��Left=�66   $��	340,5,0,0q�6Q   $	��Top=�6d   $	��100��6|   $	��100��6�   ��.  |�.  $��0q�6�   L . $7 CrystalDiskInfo��6�   ���#E0E0E0���#E0E0E0��6)  Z� .��5�     �� $#��Resources/CrystalDiskInfo.png��6    $��65��6G   $��Center=�6R   $��Center=�6o   $��65��6�   5�     �� 5�     -%   5        $9 CrystalDiskMark��5�     $
��Left=�66   $��	470,5,0,0q�6Q   $	��Top=�6d   $	��100��6|   $	��100��6�   ��.  |�.  $��0q�6�   L . $7 CrystalDiskMark��6�   ���#E0E0E0���#E0E0E0��6)  Z� .��5�     �� $#��Resources/CrystalDiskMark.png��6    $��65��6G   $��Center=�6R   $��Center=�6o   $��65��6�   5�     �� 5�     -&   5        $9 HD_Tune��5�     $
��Left=�6.   $��	600,5,0,0q�6I   $	��Top=�6\   $	��100��6t   $	��100��6�   ��.  |�.  $��0q�6�   L . $7 HD Tune��6�   ���#E0E0E0���#E0E0E0��6  Z� .��5�     �� $��Resources/HD Tune.png��6    $��65��6?   $��Center=�6J   $��Center=�6g   $��65��6�   5�     �� 5�     -'   5        $9 SSD_Z��5�     $
��Left=�6,   $��	730,5,0,0q�6G   $	��Top=�6Z   $	��100��6r   $	��100��6~   ��.  |�.  $��0q�6�   L . $7 SSD-Z��6�   ���#E0E0E0���#E0E0E0��6  Z� .��5�     �� $��Resources/SSD-Z.png��6    $��65��6=   $��Center=�6H   $��Center=�6e   $��65��6�   5�     �� 5�     -(   5        $9 
PartAssist��5�     $
��Left=�61   $��210,157.5,0,0q�6L   $	��Top=�6c   $	��100��6{   $	��100��6�   ��.  |�.  $��0q�6�   L . $7 AOMEI Partition Assistant��6�   ���#E0E0E0���#E0E0E0��62  Z� .��5�     �� $��Resources/PartAssist.png��6    $��65��6B   $��Center=�6M   $��Center=�6j   $��65��6�   5�     �� 5�     -)   5        $9 Disk_Benchmark��5�     $
��Left=�65   $��	860,5,0,0q�6P   $	��Top=�6c   $	��100��6{   $	��100��6�   ��.  |�.  $��0q�6�   L . $7 ATTO Disk Benchmark��6�   ���#E0E0E0���#E0E0E0��6,  Z� .��5�     �� $"��Resources/Disk Benchmark.png��6    $��80��6F   $��Center=�6Q   $��Center=�6n   $��80��6�   5�     �� 5�     -*   5        $9 
DiskGenius��5�     $
��Left=�61   $��340,157.5,0,0q�6L   $	��Top=�6c   $	��100��6{   $	��100��6�   ��.  |�.  $��0q�6�   L . $7 
DiskGenius��6�   ���#E0E0E0���#E0E0E0��6#  Z� .��5�     �� $��Resources/DiskGenius.png��6    $��65��6B   $��Center=�6M   $��Center=�6j   $��65��6�   5�       5�     -+   5        $? HDRemoveCard��5�  *   $
��Left=�6@   $	��Top=�6[   $	��150��6s   $	��150��6�   ��.  $/7 )硬盘工具(该列为格式化专区等)��6�   $��	30,10,0,0q�6�   $��2��6�   .��5�     � $��Center=�6   $��0q�6<   $��Center=�6G   $	��130��6b   $	��130��6n   .+�5�       $: HarddiskRemove=�66   $	��100��6L   $	��100��6Y   ���#212121$��Center=�6z   $��Center=�6�   5�     5�     �� 5�     -,   5        $9 Victoria��5�     $
��Left=�6/   $	��Top=�6J   $	��100��6b   $	��100��6n   ��.  |�.  $��0q�6�   L . $7 Victoria��6�   ���#E0E0E0���#E0E0E0$��2��6  $��
210,35,0,0q�6+  ��6  Z� .��5�     �� $��Resources/Victoria.png��6    $��65��6@   $��Center=�6K   $��Center=�6h   $��65��6�   5�     �� 5�     --   5        $9 LLFTOOL��5�     $
��Left=�6.   $	��Top=�6I   $	��100��6a   $	��100��6m   ��.  |�.  $��0q�6�   L . $7 LLFTOOL��6�   ���#E0E0E0���#E0E0E0$��2��6  $��
340,35,0,0q�6)  ��6  Z� .��5�     �� $��Resources/LLFTOOL.png��6    $��65��6?   $��Center=�6J   $��Center=�6g   $��65��6�   5�       5�     -.   5        $? Computer_Test��5�  *   $
��Left=�6A   $	��Top=�6\   $	��150��6t   $	��150��6�   ��.  $$7 电脑详细信息检测工具��6�   $��	30,10,0,0q�6�   $��4��6�   .��5�     � $��Center=�6   $��0q�6<   $��Center=�6G   $	��130��6b   $	��130��6n   .+�5�       $: Laptop=�66   $��90��6D   $��90��6P   ���#212121$��Center=�6p   $��Center=�6�   5�     5�     �� 5�     -/   5        $9 Aida64��5�     $
��Left=�6-   $	��Top=�6H   $	��100��6`   $	��100��6l   ��.  |�.  $��0q�6�   L . $7 Aida64��6�   ���#E0E0E0���#E0E0E0$��4��6  $��
210,35,0,0q�6'  ��6  Z� .��5�     �� $��Resources/Aida64.png��6    $��65��6>   $��Center=�6I   $��Center=�6f   $��65��6�   5�     �� 5�     -0   5        $9 CPU_Z��5�     $
��Left=�6,   $	��Top=�6G   $	��100��6_   $	��100��6k   ��.  |�.  $��0q�6�   L . $7 CPU-Z��6�   ���#E0E0E0���#E0E0E0$��4��6  $��
340,35,0,0q�6%  ��6  Z� .��5�     �� $��Resources/CPU-Z.png��6    $��65��6=   $��Center=�6H   $��Center=�6e   $��65��6�   5�     �� 5�     -1   5        $9 GPU_Z��5�     $
��Left=�6,   $	��Top=�6G   $	��100��6_   $	��100��6k   ��.  |�.  $��0q�6�   L . $7 GPU-Z��6�   ���#E0E0E0���#E0E0E0$��4��6  $��
470,35,0,0q�6%  ��6  Z� .��5�     �� $��Resources/GPU-Z.png��6    $��65��6=   $��Center=�6H   $��Center=�6e   $��65��6�   5�     �� 5�     -2   5        $9 HWMonitorPro��5�     $
��Left=�63   $	��Top=�6N   $	��100��6f   $	��100��6r   ��.  |�.  $��0q�6�   L . $7 HWMonitorPro��6�   ���#E0E0E0���#E0E0E0$��4��6&  $��
600,35,0,0q�63  ��6  Z� .��5�     �� $ ��Resources/HWMonitorPro.png��6    $��65��6D   $��Center=�6O   $��Center=�6l   $��65��6�   5�     �� 5�     -3   5        $9 HWiNFO��5�     $
��Left=�6-   $	��Top=�6H   $	��100��6`   $	��100��6l   ��.  |�.  $��0q�6�   L . $7 HWiNFO��6�   ���#E0E0E0���#E0E0E0$��4��6  $��
730,35,0,0q�6'  ��6  Z� .��5�     �� $��Resources/HWiNFO.png��6    $��65��6>   $��Center=�6I   $��Center=�6f   $��65��6�   5�     �� 5�     -4   5        $9 DisplayX��5�     $
��Left=�6/   $	��Top=�6J   $	��100��6b   $	��100��6n   ��.  |�.  $��0q�6�   L . $7 DisplayX��6�   ���#E0E0E0���#E0E0E0$��4��6  $��
860,35,0,0q�6+  ��6  Z� .��5�     �� $��Resources/DisplayX.png��6    $��65��6@   $��Center=�6K   $��Center=�6h   $��65��6�   5�     5�     5�     	  5�     -5   5        $C AppSnackbar��5�  &   D .  $��	0,-62,0,0q�6L   $��0q�6_   $��Bottom=�6k   .E 5�     
  -6   5        $F AppSnackbarMessage��5�  1   $S��M错误！ 无法打开软件 请检查根目录文件夹内软件是否存在��6M   $G 好的��6s   |�.  ��.  $��Center=�6�   $��Center=�6�   5�     5�     5�     5�     