   M S B A M L   `   `   `  ����U  QCustom Report Debugger 2.0, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullI)clr-namespace:Custom_Report_Debugger_2._0Custom_Report_Debugger_2._0  5   	   ,    &Custom_Report_Debugger_2._0.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ;local)clr-namespace:Custom_Report_Debugger_2._0     5   	     � Title$  Custom Report Debugger 2��5   	   $	��622��6*   $	��800��67   .��5	      � $
��Left=�6   $	��591��6&   $	��Top=�63   $	��790��6K   .+�5
   
   � -   5         �Name$ selectAndLoad_Grid��5
      .+�5      �� -   5         ��Name$ 
textToTest��5      $
��Left=�6.   $	��282��6I   $��
27,179,0,0q�6V   $	��Top=�6j   $	��740��6�    � Cursor$ Arrow��6�    �� 
IsReadOnly .  �� 
LineHeight$ 3��6�    � ToolTip$X RThis does not live update, clicking confirm will reload the document automatically��6�   . �5      /� .T�5      J� $ 3��6    .*�6/   .*�5      �� $� ��6   5      5      5      �� 5      -   5         ��Name$ documentSelectionGuidance��5   �   $;��5Please select a document in the .docx format to test.��6   $��
27,112,0,0q�6T   $	��Top=�6h   $��43��6�   $��Center=�6�   $
��Left=�6�   $	��740��6�   �� 5      -   5         ��Name$ SelectFile_Button��5      $��Select File��61   $	��Top=�6G   $��
295,33,0,0q�6_   $��74��6s   $
��Left=�6�   $	��202��6�   �� 5      -   5        	 ��Name$	 previewDescription��5      $
��Left=�64   $��
27,155,0,0q�6O   $
��Wrap=�6c   $(��"Text taken from the selected file:��6w   $	��Top=�6�   $	��740��6�   $y�Hidden=�6�   $��14(�6�   $
��Bold"�6�   �� 5      -   5        $	 formatting_description��5      $
��Left=�69   $��
27,466,0,0q�6T   $
��Wrap=�6h   $����Please review the above to ensure it is of the correct file, each MERGEFIELD should be on a new line and include curly braces either side. Any that appear incorrectly here may not be tested correctly with this program, and should be checked manually to validate them. Once you are happy with the data above, press the confirm button below.��6|   $	��Top=�6�  $��53��6�  $	��740��6�  $y�Hidden=�6  �� 5      -   5        $ confirmation_button��5      $��Confirm��63   $��38��6E   $��327,533,0,0q�6Q   $	��Top=�6f   $
��Left=�6~   $	��137��6�   $y�Hidden=�6�   5      � 5   
   -   5        $ progress_Grid��5      $
��Left=�6&   $	��591��6A   $	��Top=�6N   $	��790��6f   $y�Hidden=�6r   .+�5      �� -	   5        $��processingStatus_Text��5      $
��Left=�68   $	��200��6S   $��
10,381,0,0q�6`   $	��Top=�6t   $	��770��6�    . $ 3��6�   &   System.Windows.Media.TextOptions
   TextFormattingMode$
 Display=�6�   . �5      /� .T�5      J� $ 3��6    .*�6/   .*�5      �� $� ��6   5      5      5       � 5!      -
   5        $��progress_bar��5!   y   $
��Left=�6   $��35��65   $��
10,341,0,0q�6A   $	��Top=�6U   $	��770��6m   6y   �� 5"      -   5        $��softwareRunning_Label��5"      $����Software running, please be patient, this should take 1 - 2 minutes... 
Whilst you wait, please enjoy some pictures of cute pets.��62   $
��Left=�6�   $��	10,10,0,0q�6�   $	��Top=�6�   $��62��6  $	��770��6  $��18(�6%  �� 5#      -   5        $��	doggo_pic��5#      $��10,72,568,255q�6&   $G��Aaa526cfc99c5f060d25fbd572e06bec9--buzzfeed-animals-silly-dogs.jpg��6=   $
��Fill=�6�   �� 5$      -   5        $��	kitty_pic��5$      $��227,72,227,255q�6&   $��cute-cat-videos-lede.jpg��6>   $
��Fill=�6`   �� 5%      -   5        $��hamster_pic��5%      $��568,72,10,255q�6(   $9��3097b43a8cf3621a039f5f5175d57362f--cute-hamsters.jpg��6?   $
��Fill=�6|   �� 5&      -   5        $��viewResults_Button��5&      $��Continue��60   $
��Left=�6C   $��323,169,0,0q�6^   $	��Top=�6s   $	��157��6�   $��65��6�   $y�Hidden=�6�   5'      � 5(   
   -   5        $ results_Grid��5(      $
��Left=�6%   $	��571��6@   $��	10,10,0,0q�6M   $	��Top=�6`   $	��770��6x   $y�Hidden=�6�   .+�5)      �� -   5        $��results_text��5)      $
��Left=�6/   $	��504��6J   $��	10,57,0,0q�6W   $	��Top=�6j   $	��750��6�    . $ 3��6�   . �5*      /� .T�5+      J� .*�6    .*�5,      �� $���Below are the full results of the various tests, please work section by section when fixing these, and re-run the checks in between. Sometimes errors found higher up in this list, will affect those further down. If you can't work out what the problem is, please speak with other members of the team to try and work it out. Please remember that this debugger is not an auto solver, it covers most issues that arise, however it is not perfect and may well miss errors, for example, it does not check that all groups opened inside tables are closed in the same table. However, if you identify something that it has missed, please let Tom know so that he can try to fix it.��6   5-      5.      5/      �� 50      $��The results are in!��6   $
��Left=�62   $��	10,10,0,0q�6M   $	��Top=�6`   $��42��6x   $	��622��6�   $��18(�6�   �� 51      -   5        $��Restart��51      $
��Left=�6'   $��
649,10,0,0q�6B   $	��Top=�6V   $	��111��6n   $��42��6z   6�   52      53      54      