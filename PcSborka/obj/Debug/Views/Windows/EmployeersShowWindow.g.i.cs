﻿#pragma checksum "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "933FCC794F5A65C3045C00F17F537A7AECA7478E7741A38BD56B50D136F2D222"
//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.42000
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

using PcSborka.Views.Windows;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace PcSborka.Views.Windows {
    
    
    /// <summary>
    /// EmployeersShowWindow
    /// </summary>
    public partial class EmployeersShowWindow : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 14 "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListView Emploeeyrs_listView;
        
        #line default
        #line hidden
        
        
        #line 20 "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button Back_button;
        
        #line default
        #line hidden
        
        
        #line 21 "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button CreateEmployeer_button;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/PcSborka;component/views/windows/employeersshowwindow.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.Emploeeyrs_listView = ((System.Windows.Controls.ListView)(target));
            
            #line 14 "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml"
            this.Emploeeyrs_listView.MouseDoubleClick += new System.Windows.Input.MouseButtonEventHandler(this.Emploeeyrs_listView_MouseDoubleClick);
            
            #line default
            #line hidden
            return;
            case 2:
            this.Back_button = ((System.Windows.Controls.Button)(target));
            
            #line 20 "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml"
            this.Back_button.Click += new System.Windows.RoutedEventHandler(this.Back_button_Click);
            
            #line default
            #line hidden
            return;
            case 3:
            this.CreateEmployeer_button = ((System.Windows.Controls.Button)(target));
            
            #line 21 "..\..\..\..\Views\Windows\EmployeersShowWindow.xaml"
            this.CreateEmployeer_button.Click += new System.Windows.RoutedEventHandler(this.CreateEmployeer_button_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}
