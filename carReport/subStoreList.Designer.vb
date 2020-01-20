<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class subStoreList
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(subStoreList))
        Me.lbStore = New System.Windows.Forms.Label()
        Me.txStore = New System.Windows.Forms.TextBox()
        Me.txEtc = New System.Windows.Forms.TextBox()
        Me.lbEtc = New System.Windows.Forms.Label()
        Me.subGrid = New System.Windows.Forms.DataGridView()
        Me.subStore = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.subAddress = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.subGrid1 = New System.Windows.Forms.DataGridView()
        Me.subStore1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.subAddress1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnInput = New System.Windows.Forms.Button()
        CType(Me.subGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.subGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbStore
        '
        Me.lbStore.AutoSize = True
        Me.lbStore.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.lbStore.Location = New System.Drawing.Point(12, 15)
        Me.lbStore.Name = "lbStore"
        Me.lbStore.Size = New System.Drawing.Size(66, 19)
        Me.lbStore.TabIndex = 0
        Me.lbStore.Text = "주유소"
        '
        'txStore
        '
        Me.txStore.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.txStore.Location = New System.Drawing.Point(84, 12)
        Me.txStore.Name = "txStore"
        Me.txStore.Size = New System.Drawing.Size(204, 29)
        Me.txStore.TabIndex = 1
        '
        'txEtc
        '
        Me.txEtc.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.txEtc.Location = New System.Drawing.Point(373, 12)
        Me.txEtc.Name = "txEtc"
        Me.txEtc.Size = New System.Drawing.Size(384, 29)
        Me.txEtc.TabIndex = 3
        '
        'lbEtc
        '
        Me.lbEtc.AutoSize = True
        Me.lbEtc.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.lbEtc.Location = New System.Drawing.Point(314, 15)
        Me.lbEtc.Name = "lbEtc"
        Me.lbEtc.Size = New System.Drawing.Size(53, 19)
        Me.lbEtc.TabIndex = 2
        Me.lbEtc.Text = "주 소"
        '
        'subGrid
        '
        Me.subGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.subGrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.subStore, Me.subAddress})
        Me.subGrid.Location = New System.Drawing.Point(10, 50)
        Me.subGrid.Name = "subGrid"
        Me.subGrid.RowHeadersVisible = False
        Me.subGrid.RowTemplate.Height = 23
        Me.subGrid.Size = New System.Drawing.Size(830, 561)
        Me.subGrid.TabIndex = 4
        '
        'subStore
        '
        Me.subStore.HeaderText = "주 유 소"
        Me.subStore.Name = "subStore"
        Me.subStore.Width = 250
        '
        'subAddress
        '
        Me.subAddress.HeaderText = "주   소"
        Me.subAddress.Name = "subAddress"
        Me.subAddress.Width = 550
        '
        'subGrid1
        '
        Me.subGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.subGrid1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.subStore1, Me.subAddress1})
        Me.subGrid1.Location = New System.Drawing.Point(10, 50)
        Me.subGrid1.Name = "subGrid1"
        Me.subGrid1.RowHeadersVisible = False
        Me.subGrid1.RowTemplate.Height = 23
        Me.subGrid1.Size = New System.Drawing.Size(830, 561)
        Me.subGrid1.TabIndex = 5
        Me.subGrid1.Visible = False
        '
        'subStore1
        '
        Me.subStore1.HeaderText = "주 유 소"
        Me.subStore1.Name = "subStore1"
        Me.subStore1.Width = 250
        '
        'subAddress1
        '
        Me.subAddress1.HeaderText = "주   소"
        Me.subAddress1.Name = "subAddress1"
        Me.subAddress1.Width = 550
        '
        'btnInput
        '
        Me.btnInput.Location = New System.Drawing.Point(772, 10)
        Me.btnInput.Name = "btnInput"
        Me.btnInput.Size = New System.Drawing.Size(68, 31)
        Me.btnInput.TabIndex = 13
        Me.btnInput.Text = "입 력"
        Me.btnInput.UseVisualStyleBackColor = True
        '
        'subStoreList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(849, 620)
        Me.Controls.Add(Me.btnInput)
        Me.Controls.Add(Me.subGrid1)
        Me.Controls.Add(Me.subGrid)
        Me.Controls.Add(Me.txEtc)
        Me.Controls.Add(Me.lbEtc)
        Me.Controls.Add(Me.txStore)
        Me.Controls.Add(Me.lbStore)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "subStoreList"
        Me.Text = "주유소와 주소 찾기"
        CType(Me.subGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.subGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbStore As System.Windows.Forms.Label
    Friend WithEvents txStore As System.Windows.Forms.TextBox
    Friend WithEvents txEtc As System.Windows.Forms.TextBox
    Friend WithEvents lbEtc As System.Windows.Forms.Label
    Friend WithEvents subGrid As System.Windows.Forms.DataGridView
    Friend WithEvents subStore As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents subAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents subGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents subStore1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents subAddress1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnInput As System.Windows.Forms.Button
End Class
