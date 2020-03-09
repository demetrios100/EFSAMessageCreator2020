// <copyright file="App.xaml.cs" company="EFSAUsersGroup">Copyright (c) EFSA Users Group. All rights reserved.</copyright>
// <author>Demetrios Ioannides</author>
// <email>dvi1@columbia.edu</email>
// <summary>The App xaml for the EFSA Message Creator Utility</summary>

namespace EFSAMessageCreator
{
    using System;
    using System.Windows;

    /// <summary>
    /// Interaction logic for App XAML
    /// </summary>
    public partial class App : Application
    {
        // Variables to be changed when a new Schema is used      
#if Debug
            public static String Schema = "CHEM_MON_2020.xsd";
            public static String ElementMappingFileName = "ElementMapping.xml";
            public static String OutputXMLFileName = "Output.xml";
            public static String ApplicationTitle = "EFSA Message Creator 2020";
            public static String ApplicationIcon = "Chemicals.ico";
            public static String ApplicationBackground = "#98deeb";
#elif Release
            public static String Schema = "CHEM_MON_2020.xsd";
            public static String ElementMappingFileName = "ElementMapping.xml";
            public static String OutputXMLFileName = "Output.xml";
            public static String ApplicationTitle = "EFSA Message Creator 2020";
            public static String ApplicationIcon = "Chemicals.ico";
            public static String ApplicationBackground = "#98deeb";
#endif
    }
}
