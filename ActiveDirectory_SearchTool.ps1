# Working progress of a script which will allow some basic AD querying 
# and modifying via a windows.form  GUI.



# create main form.
Add-Type -assembly System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text ='GUI AD Search Tool by Andrew Faulds (V1.0 beta)'
$form.Width = 600
$form.Height = 400
$form.AutoSize = $true



# Create label field for the search
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Search AD User (surname): "
$Label.AutoSize = $true
#$form.Controls.Add($Label)


# label for second tab
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Search AD Group (exact name): "
$Label2.AutoSize = $true
#$form.Controls.Add($Label2)


# label for third tab
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = "Search for user's member groups (by UUN): "
$Label3.AutoSize = $true
#$form.Controls.Add($Label3)


# Create textbox to enter search string into
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
#$form.Controls.Add($textBox)

# Create second textbox to enter search string into (tab 2)
$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,40)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
#$form.Controls.Add($textBox2)


# Create third textbox to enter search string into (tab 3)
$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(10,40)
$textBox3.Size = New-Object System.Drawing.Size(260,20)
#$form.Controls.Add($textBox3)


# Create button to initiate query.(tab 1)
$Search_Button = New-Object System.Windows.Forms.Button
$Search_Button.Location = New-Object System.Drawing.Size(170,70)
$Search_Button.Size = New-Object System.Drawing.Size(120,23)
$Search_Button.Text = "Check AD"

# Create action button for second tab

$Search_Button2 = New-Object System.Windows.Forms.Button
$Search_Button2.Location = New-Object System.Drawing.Size(170,70)
$Search_Button2.Size = New-Object System.Drawing.Size(120,23)
$Search_Button2.Text = "Search AD"

# Create action button for third tab

$Search_Button3 = New-Object System.Windows.Forms.Button
$Search_Button3.Location = New-Object System.Drawing.Size(170,70)
$Search_Button3.Size = New-Object System.Drawing.Size(120,23)
$Search_Button3.Text = "Search AD"



# Create exit button (tab 1)
$button_exit = New-Object System.Windows.Forms.Button
$button_exit.Text = "Exit"
$button_exit.Location = New-Object System.Drawing.Point(170,480)


# Create exit button (tab 2)
$button_exit2 = New-Object System.Windows.Forms.Button
$button_exit2.Text = "Exit"
$button_exit2.Location = New-Object System.Drawing.Point(170,480)


# Create exit button (tab 3)
$button_exit3 = New-Object System.Windows.Forms.Button
$button_exit3.Text = "Exit"
$button_exit3.Location = New-Object System.Drawing.Point(170,480)




# create grid to return search results in tab 1
$ad_grid = New-Object Windows.Forms.DataGridview
$ad_grid.DataBindings.DefaultDataSourceUpdateMode = 0
$ad_grid.Name = "grouplist"
$ad_grid.DataMember = ""
$ad_grid.TabIndex = 1
$ad_grid.Location = New-Object Drawing.Point 300,40
$ad_grid.Size = New-Object Drawing.Point 800,440
$ad_grid.readonly = $true
$ad_grid.AutoSizeColumnsMode = 'AllCells'
$ad_grid.SelectionMode = 'FullRowSelect'
$ad_grid.MultiSelect = $false
$ad_grid.RowHeadersVisible = $false
$ad_grid.allowusertoordercolumns = $true

# second grid 
$ad_grid2 = New-Object Windows.Forms.DataGridview
$ad_grid2.DataBindings.DefaultDataSourceUpdateMode = 0
$ad_grid2.Name = "grouplist2"
$ad_grid2.DataMember = ""
$ad_grid2.TabIndex = 2
$ad_grid2.Location = New-Object Drawing.Point 300,40
$ad_grid2.Size = New-Object Drawing.Point 800,440
$ad_grid2.readonly = $true
$ad_grid2.AutoSizeColumnsMode = 'AllCells'
$ad_grid2.SelectionMode = 'FullRowSelect'
$ad_grid2.MultiSelect = $false
$ad_grid2.RowHeadersVisible = $false
$ad_grid2.allowusertoordercolumns = $true

# Third Grid
$ad_grid3 = New-Object Windows.Forms.DataGridview
$ad_grid3.DataBindings.DefaultDataSourceUpdateMode = 0
$ad_grid3.Name = "grouplist3"
$ad_grid3.DataMember = ""
$ad_grid3.TabIndex = 3
$ad_grid3.Location = New-Object Drawing.Point 300,40
$ad_grid3.Size = New-Object Drawing.Point 800,440
$ad_grid3.readonly = $true
$ad_grid3.AutoSizeColumnsMode = 'AllCells'
$ad_grid3.SelectionMode = 'FullRowSelect'
$ad_grid3.MultiSelect = $false
$ad_grid3.RowHeadersVisible = $false
$ad_grid3.allowusertoordercolumns = $true






# Create the tabcontrols
$tabcontrol = New-Object windows.Forms.TabControl
$tabpage_ad = New-Object windows.Forms.TabPage
$tabpage_group = New-Object windows.Forms.TabPage
$tabpage_ugroup = New-Object windows.Forms.Tabpage
$tabcontrol.width = 1150
$tabcontrol.Height = 540
 

#Tab Controls 1
$tabcontrol.tabpages.add($tabpage_ad)
$tabpage_ad.controls.add($Label)
$tabpage_ad.controls.add($textBox)
$tabpage_ad.controls.add($Search_Button)
$tabpage_ad.controls.add($ad_grid)
$tabpage_ad.controls.add($button_exit)

#Tab Controls 2

$tabcontrol.tabpages.add($tabpage_group)
$tabpage_group.controls.add($textBox2)
$tabpage_group.controls.add($ad_grid2)
$tabpage_group.controls.add($Label2)
$tabpage_group.controls.add($Search_Button2)
$tabpage_group.controls.add($button_exit2)


# Tab Controls 3 

$tabcontrol.tabpages.add($tabpage_ugroup)
$tabpage_ugroup.controls.add($textBox3)
$tabpage_ugroup.controls.add($ad_grid3)
$tabpage_ugroup.controls.add($Label3)
$tabpage_ugroup.controls.add($Search_Button3)
$tabpage_ugroup.controls.add($button_exit3)




#Set TabPages text
$tabpage_ad.Text = "AD User Info"
$tabpage_group.Text = "AD Group Members"
$tabpage_ugroup.Text = "AD User's Member Groups"

# Add the tab controls to the Form

$form.Controls.Add($tabcontrol)


#################tab 1 event handlers  ##################################

# Add click functiont to button and...
# Set up event handler to extarct text from TextBox and display it on the Label.
$Search_Button.add_click({
$grid = $ad_grid
$query = "*"+$textBox.Text.ToString()+"*"
search_contact_ad($query)
})

# same as above but allows you to use the return key instead of clicking.
$textBox.add_KeyPress({
If ($_.KeyChar -eq 13) {
$grid = $ad_grid
$query = "*"+$textBox.Text.ToString()+"*"
search_contact_ad($query)
}
})

# adds functionaility to the exit button. 

$button_exit.add_click({
$form.Close()
})
 
####################### tab 2 event handlers #####################################


$Search_Button2.add_click({
$grid2 = $ad_grid2
$query2 = $textBox2.Text.ToString()
search_group_ad($query2)
})

# same as above but allows you to use the return key instead of clicking.
$textBox2.add_KeyPress({
If ($_.KeyChar -eq 13) {
$grid2 = $ad_grid2
$query2 = $textBox2.Text.ToString()
search_group_ad($query2)
}
})

# adds functionaility to the exit button. 

$button_exit2.add_click({
$form.Close()
})


####################### tab 3 event handlers #####################################


$Search_Button3.add_click({
$grid3 = $ad_grid3
$query3 = $textBox3.Text.ToString()
search_group_mems($query3)
})

# same as above but allows you to use the return key instead of clicking.
$textBox3.add_KeyPress({
If ($_.KeyChar -eq 13) {
$grid3 = $ad_grid3
$query3 = $textBox3.Text.ToString()
search_group_mems($query3)
}
})

# adds functionaility to the exit button. 

$button_exit3.add_click({
$form.Close()
})











################ fucntions ##############################


# decleration of the actual ad search function
# write output to the view grid
# refreshes the form. (unnecessary, form seems to be event driven)

# catch blank entry exception with an error (not working atm)
#### tab 1 function ###########

function search_contact_ad([string]$lastname){
if ($lastname) {
$array_ad = New-Object System.Collections.ArrayList 
$Script:procInfo = @(Get-ADUser -Filter {Surname -like $lastname} | Get-Adobject -properties * |  Select-Object @{Name="UUN";Expression={$_.Name}},@{Name="Surname";Expression={$_.sn}},@{Name="Display Name";Expression={$_.Displayname}},@{Name="Department";Expression={$_.Department}},@{Name="Title";Expression={$_.Title}},@{Name="Email";Expression={$_.userPrincipalName}} | sort-object -property UUN)
$array_ad.AddRange($procInfo)
$grid.DataSource = $array_ad
#$form.refresh()
}
else {
[windows.forms.messagebox]::show('Please enter a value',[Windows.Forms.MessageBoxIcon]::Warning)
}
}

######## tab 2 function ##########
function search_group_ad([string]$group){
if ($group) {
$array_group = New-Object System.Collections.ArrayList 
#$Script:procInfo2 = @(Get-ADgroupmember $group | sort-object -property Name | Select-Object @{Name="UUN";Expression={$_.Name}})

$Script:procInfo2 = @(Get-ADGroupMember -Identity $group | Get-ADObject -Properties * | Select-Object Displayname, Name, Title, Department | sort-object -property Displayname | Select-Object @{Name="Name";Expression={$_.Displayname}},@{Name="UUN";Expression={$_.Name}},@{Name="Title";Expression={$_.Title}},@{Name="Department";Expression={$_.Department}})
$array_group.AddRange($procInfo2)
$grid2.DataSource = $array_group
#$form.refresh()
}
else {
[windows.forms.messagebox]::show('Please enter a value',[Windows.Forms.MessageBoxIcon]::Warning)
}
}


######## tab 3 function ##########
function search_group_mems([string]$user){
if ($user) {
$array_groupmem = New-Object System.Collections.ArrayList 
$Script:procInfo3 = @(Get-ADPrincipalGroupMembership $user| Select-Object name | sort-object -property name | Select-Object @{Name="Group-Name";Expression={$_.Name}})
$array_groupmem.AddRange($procInfo3)
$grid3.DataSource = $array_groupmem
#$form.refresh()
}
else {
[windows.forms.messagebox]::show('Please enter a value',[Windows.Forms.MessageBoxIcon]::Warning)
}
}


# present form to the screen
$form.ShowDialog()

