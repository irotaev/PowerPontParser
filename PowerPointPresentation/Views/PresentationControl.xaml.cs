﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PowerPointPresentation.Views
{
  /// <summary>
  /// Interaction logic for PresentationControl.xaml
  /// </summary>
  public partial class PresentationControl : UserControl
  {
    public PresentationControl()
    {
      InitializeComponent();
    }

    private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
    {
      ((Panel)this.Parent).Children.Remove(this);
    }
  }
}
