﻿//------------------------------------------------------------------------------
// <auto-generated>
//     此代码由工具生成。
//     运行时版本:4.0.30319.42000
//
//     对此文件的更改可能会导致不正确的行为，并且如果
//     重新生成代码，这些更改将会丢失。
// </auto-generated>
//------------------------------------------------------------------------------

namespace excelXLL {
    using System;
    
    
    /// <summary>
    ///   一个强类型的资源类，用于查找本地化的字符串等。
    /// </summary>
    // 此类是由 StronglyTypedResourceBuilder
    // 类通过类似于 ResGen 或 Visual Studio 的工具自动生成的。
    // 若要添加或移除成员，请编辑 .ResX 文件，然后重新运行 ResGen
    // (以 /str 作为命令选项)，或重新生成 VS 项目。
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resource1 {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resource1() {
        }
        
        /// <summary>
        ///   返回此类使用的缓存的 ResourceManager 实例。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("excelXLL.Resource1", typeof(Resource1).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   使用此强类型资源类，为所有资源查找
        ///   重写当前线程的 CurrentUICulture 属性。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;UTF-8&quot;?&gt;
        ///&lt;!--中华人秘共和国行政区划代码即身份证前6位编码（截止2016年7月31日）--&gt;
        ///&lt;!--行政区划代码查询网址：http://www.stats.gov.cn/tjsj/tjbz/xzqhdm/--&gt;
        ///&lt;China&gt;
        ///  &lt;provinces id=&quot;110000&quot; name=&quot;北京市&quot;&gt;
        ///    &lt;city id=&quot;110100&quot; name=&quot;市辖区&quot;&gt;
        ///      &lt;count id=&quot;110101&quot; name=&quot;东城区&quot; /&gt;
        ///      &lt;count id=&quot;110102&quot; name=&quot;西城区&quot; /&gt;
        ///      &lt;count id=&quot;110105&quot; name=&quot;朝阳区&quot; /&gt;
        ///      &lt;count id=&quot;110106&quot; name=&quot;丰台区&quot; /&gt;
        ///      &lt;count id=&quot;110107&quot; name=&quot;石景山区&quot; /&gt;
        ///      &lt;count id=&quot;110108&quot; name=&quot;海淀区&quot; /&gt;
        ///      &lt;count id=&quot;110109&quot; name=&quot;门头沟区&quot; /&gt;
        ///   [字符串的其余部分被截断]&quot;; 的本地化字符串。
        /// </summary>
        internal static string AreaNumber {
            get {
                return ResourceManager.GetString("AreaNumber", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找 System.Drawing.Bitmap 类型的本地化资源。
        /// </summary>
        internal static System.Drawing.Bitmap img_star {
            get {
                object obj = ResourceManager.GetObject("img_star", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查找 System.Drawing.Bitmap 类型的本地化资源。
        /// </summary>
        internal static System.Drawing.Bitmap img_star_blue {
            get {
                object obj = ResourceManager.GetObject("img_star_blue", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查找 System.Drawing.Bitmap 类型的本地化资源。
        /// </summary>
        internal static System.Drawing.Bitmap img_star_red {
            get {
                object obj = ResourceManager.GetObject("img_star_red", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
    }
}
