namespace AutoReconciliation.Models
{
    // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false, ElementName = "BOM")]

    public partial class BusinessPartners
    {

        private BOMBO boField;

        /// <remarks/>
        public BOMBO BO
        {
            get
            {
                return this.boField;
            }
            set
            {
                this.boField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class BOMBO
    {

        private BOMBOAdmInfo admInfoField;

        private BOMBORow[] oCRDField;

        /// <remarks/>
        public BOMBOAdmInfo AdmInfo
        {
            get
            {
                return this.admInfoField;
            }
            set
            {
                this.admInfoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("row", IsNullable = false)]
        public BOMBORow[] OCRD
        {
            get
            {
                return this.oCRDField;
            }
            set
            {
                this.oCRDField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class BOMBOAdmInfo
    {

        private sbyte objectField;

        /// <remarks/>
        public sbyte Object
        {
            get
            {
                return this.objectField;
            }
            set
            {
                this.objectField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class BOMBORow
    {

        private string cardCodeField;

        /// <remarks/>
        public string CardCode
        {
            get
            {
                return this.cardCodeField;
            }
            set
            {
                this.cardCodeField = value;
            }
        }
    }
}