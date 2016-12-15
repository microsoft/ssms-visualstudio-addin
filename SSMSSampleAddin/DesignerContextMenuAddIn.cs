namespace Microsoft.Dynamics.Samples.AddIns.SSMSAddin
{
    using System;
    using System.ComponentModel.Composition;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;

    // These namespaces contain the support for the VS metadata API, the APIs
    // describing what is currently open in VS.
    using Microsoft.Dynamics.Framework.Tools.Extensibility;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Core;
    using Microsoft.Dynamics.Framework.Tools.Configuration;

    // Convenience prefixes
    using TablesAutomation = Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Tables;
    using ViewsAutomation = Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Views;

    // These namespaces contain the general purpose metadata APIs that work
    // irrespective of the state of VS.
    using Metadata = Microsoft.Dynamics.AX.Metadata;
    using Microsoft.Dynamics.AX.Metadata.MetaModel;

    /// <summary>
    /// This addin opens a window in Microsoft SQL Management studio that allows
    /// the user to query the table that the user has selected in VS.
    /// </summary>
    /// <remarks>Since the class derives from DesignerMenuBase, this addin is
    /// available from the designer. The attributes below tell VS to include this addin when the given types are 
    /// selected in the designer. 
    /// </remarks>
    [Export(typeof(IDesignerMenu))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(TablesAutomation.ITable))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(TablesAutomation.ITableExtension))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(TablesAutomation.IBaseField))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(TablesAutomation.IFieldGroup))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(TablesAutomation.ITableHasRelations))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(TablesAutomation.IRelation))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ViewsAutomation.IView))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ViewsAutomation.IViewBaseField))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ViewsAutomation.IFieldGroup))]
    public class SelectInSSMSDesignerAddIn : DesignerMenuBase
    {
        private static string BusinessDatabaseName
        {
            get
            {
                return ConfigurationHelper.CurrentConfiguration.BusinessDatabaseName;
            }
        }

        #region Properties
        /// <summary>
        /// Caption for the menu item. This is what users would see in the add-in menu.
        /// </summary>
        public override string Caption
        {
            get
            {
                return SSMSSampleAddin.AddinResources.SelectInSSMSDesignerAddinCaption;
            }
        }

        /// <summary>
        /// Unique name of the add-in
        /// </summary>
        public override string Name
        {
            get
            {
                return "Select in SQL Server Management Studio";
            }
        }

        /// <summary>
        /// Backing field for the metadata provider that is useful for retrieving
        /// metadata that is not loaded in the VS instance.
        /// </summary>
        private Metadata.Providers.IMetadataProvider metadataProvider = null;

        /// <summary>
        /// Gets a singleton instance of the metadata provider that can access the metadata repository.
        /// Any metadata, irrespective of whether it is part of what is being edited by VS, is available
        /// through this provider.
        /// </summary>
        public Metadata.Providers.IMetadataProvider MetadataProvider
        {
            get
            {
                if (this.metadataProvider == null)
                {
                    this.metadataProvider = DesignMetaModelService.Instance.CurrentMetadataProvider;
                }
                return this.metadataProvider;
            }
        }

        #endregion

        /// <summary>
        /// Resolve the given label to the text it represents in the default language. If a non 
        /// label is passed, that text is returned unchanged.
        /// </summary>
        /// <param name="label">The label to look up.</param>
        /// <returns>The text for the label resolved to the current language.</returns>
        private string ResolveLabel(string label)
        {
            var labelResolver = CoreUtility.ServiceProvider.GetService(typeof(Microsoft.Dynamics.Framework.Tools.Integration.Interfaces.ILabelResolver)) as Microsoft.Dynamics.Framework.Tools.Integration.Interfaces.ILabelResolver;

            if (labelResolver != null)
            {
                return labelResolver.GetLabelText(label);
            }
            return label;
        }

        /// <summary>
        /// Get the label provided in the field metadata. If no such label is provided, check
        /// if the label's type is an EDT, that possibly has a label. If that fails, return null.
        /// </summary>
        /// <param name="field">The metadata instance describing the field</param>
        /// <returns>The field label, if one is found, or null.</returns>
        private string GetFieldLabel(AxTableField field)
        {
            // if the field has a label, use it.
            var fieldLabel = this.ResolveLabel(field.Label);

            if (!string.IsNullOrWhiteSpace(fieldLabel))
            {
                return fieldLabel;
            }

            // if not, use the label on the EDT, if there is one
            if (!string.IsNullOrWhiteSpace(field.ExtendedDataType))
            {
                // Fetch the EDT through the metadata provider
                AxEdt edt = this.MetadataProvider.Edts.Read(field.ExtendedDataType);
                if (edt != null)
                {
                    var edtLabel = this.ResolveLabel(edt.Label);
                    if (!string.IsNullOrWhiteSpace(edtLabel))
                    {
                        return edtLabel;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get the label from the metadata describing a field on a view.
        /// </summary>
        /// <param name="field">The view field metadata</param>
        /// <returns>The label, if one is defined. Otherwise, returns null</returns>
        private string GetFieldLabel(AxViewField field)
        {
            // if the field has a label, use it.
            var fieldLabel = this.ResolveLabel(field.Label);

            if (!string.IsNullOrWhiteSpace(fieldLabel))
            {
                return fieldLabel;
            }

            return null;
        }

        #region Utility functions for adding lines to the file
        /// <summary>
        /// Append the field in the given tabular object to the output, one field per line.
        /// </summary>
        /// <param name="result">The string containing the result</param>
        /// <param name="tableName">The unmangled name of the tabular object</param>
        /// <param name="fieldName">The unmangled name of the field</param>
        /// <param name="fieldLabel">An optional label to use as the prompt. If this is null, the name is used.</param>
        /// <param name="first">Determines whether the first field has been written or not.</param>
        private void AddField(StringBuilder result, string tableName, string fieldName, string label, ref bool first)
        {
            // Fields are indented to be positioned under the select statement
            result.Append("    ");

            if (!first)
            {
                result.Append(",");
            }

            result.Append(SqlNameMangling.GetSqlTableName(tableName) + ".[" + SqlNameMangling.GetValidSqlNameForField(fieldName) + "]");
            result.AppendLine(string.Format(CultureInfo.InvariantCulture, " as '{0}'", label ?? fieldName));

            first = false;
        }

        /// <summary>
        /// Add the commaseparated fields in the collection by calling AddField on each one.
        /// If the field is an array field, all the array elements are added.
        /// </summary>
        /// <param name="result">The result where the fields are added</param>
        /// <param name="fields">The collection of fields to add</param>
        /// <param name="first">A parameter specifying whether or not this is the first field.</param>
        private void AddFields(StringBuilder result, AxTable table, IEnumerable<AxTableField> fields, ref bool first)
        {
            foreach (AxTableField field in fields)
            {
                if (field.SaveContents == Metadata.Core.MetaModel.NoYes.Yes)
                {
                    var label = this.GetFieldLabel(field);
                    this.AddField(result, table.Name, field.Name, label, ref first);

                    var edt = field.ExtendedDataType;
                    if (!string.IsNullOrWhiteSpace(edt))
                    {
                        // See if it happens to be an array field. If so, the first index
                        // does not have a suffix ([<n>]), and has already been written.
                        if (this.MetadataProvider.Edts.Exists(edt))
                        {
                            AxEdt typeDefinition = this.metadataProvider.Edts.Read(edt);
                            for (int i = 2; i <= typeDefinition.ArrayElements.Count + 1; i++)
                            {
                                var fn = field.Name + "[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                                this.AddField(result, table.Name, fn, null, ref first);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Add the system fields to the output file, if the metadata describing the table
        /// indicates that the field is used.
        /// </summary>
        /// <param name="result">The string builder that contains the output</param>
        /// <param name="table">The metadata describing the table.</param>
        /// <param name="first">A parameter specifying whether or not this is the first field.</param>
        private void AddSystemFields(StringBuilder result, AxTable table, ref bool first)
        {
            if (table.CreatedBy == Metadata.Core.MetaModel.NoYes.Yes)
            {
                this.AddField(result, table.Name, "CreatedBy", null, ref first);
            }

            if (table.CreatedDateTime == Metadata.Core.MetaModel.NoYes.Yes)
            {
                this.AddField(result, table.Name, "CreatedDateTime", null, ref first);
            }

            if (table.ModifiedBy == Metadata.Core.MetaModel.NoYes.Yes)
            {
                this.AddField(result, table.Name, "ModifiedBy", null, ref first);
            }

            if (table.ModifiedDateTime == Metadata.Core.MetaModel.NoYes.Yes)
            {
                this.AddField(result, table.Name, "ModifiedDateTime", null, ref first);
            }

            if (table.SaveDataPerPartition == Metadata.Core.MetaModel.NoYes.Yes)
            {
                this.AddField(result, table.Name, "Partition", null, ref first);
            }

            if (table.SaveDataPerCompany == Metadata.Core.MetaModel.NoYes.Yes)
            {
                this.AddField(result, table.Name, "DataAreaId", null, ref first);
            }

            this.AddField(result, table.Name, "RecId", null, ref first);
        }

        #endregion

        /// <summary>
        /// Find the sequence of tables from the table to its root in the inheritance chain.
        /// If the indicated table is not part of a supertype subtype relationship, the 
        /// result contains just the single table.
        /// </summary>
        /// <param name="leafTableName">The most derived table.</param>
        /// <returns>The sequence of tables, with the root at the top of the stack.</returns>
        private Stack<AxTable> SuperTables(string leafTableName)
        {
            Stack<AxTable> result = new Stack<AxTable>();
            AxTable table = this.MetadataProvider.Tables.Read(leafTableName);

            while (table.SupportInheritance == Metadata.Core.MetaModel.NoYes.Yes
                && !string.IsNullOrWhiteSpace(table.Extends))
            {
                result.Push(table);
                table = this.MetadataProvider.Tables.Read(table.Extends);
            }

            result.Push(table); // stack the root.
            return result;
        }


        /// <summary>
        /// The method is called when the user has selected to view a single table.
        /// The method generates a select on all fields, with no joins.
        /// </summary>
        /// <param name="selectedTable">The designer metadata designating the table.</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromTable(TablesAutomation.ITable selectedTable)
        {
            var result = new StringBuilder();

            // It is indeed a table. Look at the properties
            if (!selectedTable.IsKernelTable)
            {
                bool first = true;

                result.AppendLine(string.Format(CultureInfo.InvariantCulture, "use {0}", BusinessDatabaseName));
                result.AppendLine("go");

                Stack<AxTable> tables = this.SuperTables(selectedTable.Name);

                // List any developer documentation as a SQL comment:
                if (!string.IsNullOrEmpty(selectedTable.DeveloperDocumentation))
                {
                    result.Append("-- " + selectedTable.Name);
                    result.AppendLine(" : " + this.ResolveLabel(selectedTable.DeveloperDocumentation));
                }
                else
                {
                    result.AppendLine();
                }

                result.AppendLine("select ");

                this.AddFields(result, tables.First(), tables.First().Fields, ref first);
                this.AddSystemFields(result, tables.First(), ref first);

                result.AppendLine("from " + SqlNameMangling.GetSqlTableName(tables.First().Name));

                // If this table saves data per company, then add the where clause for
                // the user to fill out or ignore.
                if (tables.First().SaveDataPerCompany == Metadata.Core.MetaModel.NoYes.Yes)
                {
                    result.AppendLine("-- where " + SqlNameMangling.GetValidSqlNameForField("DataAreaId") + "  = 'DAT'");
                }
            }

            return result;
        }

        /// <summary>
        /// The method is called when the user has selected to view a single view.
        /// The method generates a select on all fields, with no joins.
        /// </summary>
        /// <param name="selectedTable">The designer metadata designating the view.</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromView(ViewsAutomation.IView selectedView)
        {
            var result = new StringBuilder();

            result.AppendLine(string.Format(CultureInfo.InvariantCulture, "use {0}", BusinessDatabaseName));
            result.AppendLine("go");

            if (!string.IsNullOrEmpty(selectedView.DeveloperDocumentation))
            {
                result.Append("-- " + selectedView.Name);
                result.AppendLine(" : " + this.ResolveLabel(selectedView.DeveloperDocumentation));
            }
            else
            {
                result.AppendLine();
            }

            result.AppendLine("select * ");
            result.AppendLine("from " + SqlNameMangling.GetSqlTableName(selectedView.Name));

            return result;
        }

        /// <summary>
        /// The method is called when the user has selected to view a table extension instance.
        /// The method generates a select on all fields, with no joins.
        /// </summary>
        /// <param name="selectedExtensionTable">The designer metadata designating the view.</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromTableExtension(TablesAutomation.ITableExtension selectedExtensionTable)
        {
            var result = new StringBuilder();
            bool first = true;

            result.AppendLine(string.Format(CultureInfo.InvariantCulture, "use {0}", BusinessDatabaseName));
            result.AppendLine("go");
            result.AppendLine();

            AxTableExtension extension = this.MetadataProvider.TableExtensions.Read(selectedExtensionTable.Name);
            var baseTableName = selectedExtensionTable.Name.Split('.').First();
            var tables = this.SuperTables(baseTableName);

            HashSet<string> extendedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var extendedField in extension.Fields)
            {
                extendedFields.Add(extendedField.Name);
            }

            result.AppendLine("select ");

            // List the extension fields first...
            this.AddFields(result, tables.First(), extension.Fields, ref first);

            // Then the normal ones...
            this.AddFields(result, tables.First(), tables.First().Fields.Where(f => !extendedFields.Contains(f.Name)), ref first);

            // And then system fields
            this.AddSystemFields(result, tables.First(), ref first);

            result.AppendLine("from " + SqlNameMangling.GetSqlTableName(tables.First().Name));

            if (tables.First().SaveDataPerCompany == Metadata.Core.MetaModel.NoYes.Yes)
            {
                result.AppendLine("-- where " + SqlNameMangling.GetValidSqlNameForField("DataAreaId") + " = 'DAT'");
            }

            return result;
        }

        /// <summary>
        /// Generate the SQL command selecting the given fields from the underlying table.
        /// </summary>
        /// <param name="fields">The list of fields to select</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromTableFieldList(IEnumerable<TablesAutomation.IBaseField> fields)
        {
            var result = new StringBuilder();
            bool first = true;

            result.AppendLine(string.Format(CultureInfo.InvariantCulture, "use {0}", BusinessDatabaseName));
            result.AppendLine("go");

            if (!fields.Any())
            {
                return result;
            }

            var tableName = string.Empty;
            TablesAutomation.ITable selectedTable = fields.First().Table;
            if (selectedTable != null)
            {
                tableName = selectedTable.Name;
            }
            else
            {
                var extensionTableName = fields.First().TableExtension.Name;
                tableName = extensionTableName.Split('.').First();
            }

            Stack<AxTable> tables = this.SuperTables(tableName);

            // Expand the developer documentation, if any
            if (!string.IsNullOrEmpty(tables.First().DeveloperDocumentation))
            {
                result.Append("-- " + tables.First().Name);
                result.AppendLine(" : " + this.ResolveLabel(tables.First().DeveloperDocumentation));
            }
            else
            {
                result.AppendLine();
            }

            result.AppendLine("select");

            // Calculate the union of the fields in all the tables in the hierarchy:
            var fieldUnion = new Dictionary<string, AxTableField>(StringComparer.OrdinalIgnoreCase);
            foreach (var table in tables)
            {
                foreach (var field in table.Fields)
                {
                    fieldUnion[field.Name] = field;
                }
            }

            // Now pick the fields that the user actually wants to see
            var selectedFields = new List<AxTableField>();
            foreach (var field in fields)
            {
                selectedFields.Add(fieldUnion[field.Name]);
            }

            this.AddFields(result, tables.First(), selectedFields, ref first);
            this.AddSystemFields(result, tables.First(), ref first);

            result.AppendLine("from " + SqlNameMangling.GetSqlTableName(tables.First().Name));

            return result;
        }

        /// <summary>
        /// Generate the SQL command selecting the given fields from the underlying table.
        /// </summary>
        /// <param name="fields">The list of fields to select</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromTableFieldListWithWhere(IEnumerable<TablesAutomation.IBaseField> fields)
        {
            var result = this.GenerateFromTableFieldList(fields);

            if (fields.Any())
            {
                var tables = this.SuperTables(fields.First().Table.Name);

                // If this table saves data per company, then add the where clause for
                // the user to fill out.
                if (tables.First().SaveDataPerCompany == Metadata.Core.MetaModel.NoYes.Yes)
                {
                    // Note: There is no way to get to the actual company, since the server 
                    // is not involved in this at all.
                    result.AppendLine("-- where " + SqlNameMangling.GetValidSqlNameForField("DataAreaId") + " = 'DAT'");
                }
            }

            return result;
        }

        private AxTable FindTableInEmbeddedDataSources(IEnumerable<AxQuerySimpleEmbeddedDataSource> embeddedDataSources, string dataSourceName)
        {
            foreach (AxQuerySimpleEmbeddedDataSource dataSource in embeddedDataSources)
            {
                if (string.Compare(dataSourceName, dataSourceName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return this.MetadataProvider.Tables.Read(dataSource.Table);
                }

                return FindTableInEmbeddedDataSources(dataSource.DataSources, dataSourceName);
            }

            return null;
        }

        private AxTable FindTableInDataSource(AxView view, string dataSourceName)
        {
            foreach (AxQuerySimpleRootDataSource ds in view.ViewMetadata.DataSources)
            {
                if (string.Compare(dataSourceName, ds.Name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return this.MetadataProvider.Tables.Read(ds.Table);
                }

                return FindTableInEmbeddedDataSources(ds.DataSources, dataSourceName);
            }

            return null;
        }

        /// <summary>
        /// Generate the SQL command selecting the given fields from the underlying table.
        /// </summary>
        /// <param name="fields">The list of fields to select</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromViewFieldList(IEnumerable<ViewsAutomation.IViewBaseField> fields)
        {
            var result = new StringBuilder();

            result.AppendLine(string.Format(CultureInfo.InvariantCulture, "use {0}", BusinessDatabaseName));
            result.AppendLine("go");

            if (!fields.Any())
            {
                return result;
            }

            ViewsAutomation.IView selectedView1 = fields.FirstOrDefault().View;
            AxView view = this.MetadataProvider.Views.Read(selectedView1.Name);

            // Expand the developer documentation, if any
            if (!string.IsNullOrEmpty(view.DeveloperDocumentation))
            {
                result.Append("-- " + view.Name);
                result.AppendLine(" : " + this.ResolveLabel(view.DeveloperDocumentation));
            }
            else
            {
                result.AppendLine();
            }

            result.AppendLine("select");
            bool first = true;

            foreach (ViewsAutomation.IViewField field in fields.OfType<ViewsAutomation.IViewField>())
            {
                this.AddField(result, view.Name, field.Name, null, ref first);

                // The field name refers to a name on the datasource. Find the datasource
                // and the underlying table.
                AxTable table = this.FindTableInDataSource(view, field.DataSource);
                table = this.SuperTables(table.Name).First();

                if (table != null)
                {
                    AxTableField tableField = table.Fields[field.DataField];
                    if (tableField != null)
                    {
                        var edt = tableField.ExtendedDataType;

                        if (!string.IsNullOrWhiteSpace(edt))
                        {
                            // See if it happens to be an array field. If so, the first index
                            // does not have a suffix ([<n>]), and has already been written.
                            if (this.MetadataProvider.Edts.Exists(edt))
                            {
                                AxEdt typeDefinition = this.metadataProvider.Edts.Read(edt);
                                for (int i = 2; i <= typeDefinition.ArrayElements.Count + 1; i++)
                                {
                                    var fn = field.Name + "[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                                    this.AddField(result, view.Name, fn, null, ref first);
                                }
                            }
                        }
                    }
                }
            }

            // Now deal with computed columns.
            foreach (ViewsAutomation.IViewComputedColumn computedColumn in fields.OfType<ViewsAutomation.IViewComputedColumn>())
            {
                this.AddField(result, view.Name, computedColumn.Name, null, ref first);
            }

            result.AppendLine("from " + SqlNameMangling.GetSqlTableName(view.Name));

            return result;
        }


        /// <summary>
        /// Generate the SQL command selecting the set of fields from the given field groups 
        /// from the underlying table.
        /// </summary>
        /// <param name="selectedFieldGroups">The list of field groups to select fields from.</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromTableFieldGroups(IEnumerable<TablesAutomation.IFieldGroup> selectedFieldGroups)
        {
            var result = new StringBuilder();

            if (!selectedFieldGroups.Any())
            {
                return result;
            }

            // We are implicitly assuming that all the field groups belong to the same table.
            TablesAutomation.ITable selectedTable = selectedFieldGroups.FirstOrDefault().Table;

            // Used to capture the set of fields containing the union of the fields
            // from the selected field groups.
            var fields = new HashSet<string>();

            // Unwrap the fields in all of the selected field groups, using the set
            // to weed out duplicate fields (that occur in multiple field groups)
            foreach (var fieldGroup in selectedFieldGroups)
            {
                var refs = fieldGroup.FieldGroupFieldReferences;
                foreach (IFieldGroupFieldReference reference in refs)
                {
                    fields.Add(reference.DataField);
                }
            }

            // Now the (unique) set of fields is captured.
            // Generate the SQL query from the IBaseField references.
            IList<TablesAutomation.IBaseField> fieldsToSelect = new List<TablesAutomation.IBaseField>();

            foreach (TablesAutomation.IBaseField child in selectedTable.BaseFields)
            {
                if (fields.Contains(child.Name))
                    fieldsToSelect.Add(child);
            }

            result = this.GenerateFromTableFieldList(fieldsToSelect);

            // If this table saves data per company, then add the where clause for
            // the user to fill out.
            if (selectedTable.SaveDataPerCompany == Metadata.Core.MetaModel.NoYes.Yes)
            {
                // Note: There is no way to get to the actual company, since the server 
                // is not involved in this at all.
                result.AppendLine("-- where " + SqlNameMangling.GetValidSqlNameForField("DataAreaId") + " = 'DAT'");
            }

            return result;
        }

        /// <summary>
        /// Generate the SQL command selecting the set of fields from the given field groups 
        /// from the underlying view.
        /// </summary>
        /// <param name="selectedFieldGroups">The list of field groups to select fields from.</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromViewFieldGroups(IEnumerable<ViewsAutomation.IFieldGroup> selectedFieldGroups)
        {
            var result = new StringBuilder();

            if (!selectedFieldGroups.Any())
            {
                return result;
            }

            // Used to capture the set of fields containing the union of the fields
            // from the selected field groups.
            var fields = new HashSet<string>();
            ViewsAutomation.IView selectedView = selectedFieldGroups.FirstOrDefault().View;

            foreach (var fieldGroup in selectedFieldGroups)
            {
                var refs = fieldGroup.FieldGroupFieldReferences;
                foreach (IFieldGroupFieldReference reference in refs)
                {
                    fields.Add(reference.DataField);
                }
            }

            // Now the (unique) set of fields is captured.
            // Generate the SQL query from the IBaseField references.
            IList<ViewsAutomation.IViewBaseField> fieldsToSelect = new List<ViewsAutomation.IViewBaseField>();

            foreach (ViewsAutomation.IViewBaseField child in selectedView.ViewBaseFields)
            {
                if (fields.Contains(child.Name))
                    fieldsToSelect.Add(child);
            }

            return this.GenerateFromViewFieldList(fieldsToSelect);
        }

        /// <summary>
        /// Generate the SQL command selecting the fields from the selected table, performing 
        /// joins on the tables indicated by the relations selected.
        /// </summary>
        /// <param name="selectedRelations">The list of field groups to select fields from.</param>
        /// <returns>The string containing the SQL command.</returns>
        private StringBuilder GenerateFromTableRelations(IEnumerable<TablesAutomation.IRelation> selectedRelations)
        {
            TablesAutomation.ITable table = selectedRelations.First().Table;
            Stack<AxTable> tables = this.SuperTables(table.Name);

            IList<TablesAutomation.IBaseField> fields = new List<TablesAutomation.IBaseField>();

            foreach (TablesAutomation.IBaseField field in table.BaseFields)
            {
                if (field.SaveContents == Metadata.Core.MetaModel.NoYes.Yes)
                {
                    fields.Add(field);
                }
            }

            var result = this.GenerateFromTableFieldList(fields);
            int disambiguation = 1;

            // Now add the joins. Assume that multiple relations can be
            // selected; joins will be added for all of them.
            foreach (var relation in selectedRelations)
            {
                var relatedTabularObjectName = relation.RelatedTable;
                // TODO: In principle this could be a view. Assuming table for now

                result.AppendLine(
                    string.Format(CultureInfo.InvariantCulture, "inner join {0} as t{1}", SqlNameMangling.GetSqlTableName(relatedTabularObjectName), disambiguation));
                result.Append("on ");

                bool first = true;
                foreach (TablesAutomation.IRelationConstraint constraint in relation.RelationConstraints)
                {
                    if (!first)
                        result.Append(" and ");

                    if (constraint is TablesAutomation.IRelationConstraintField)
                    {   // Table.field = RelatedTable.relatedField
                        var fieldConstraint = constraint as TablesAutomation.IRelationConstraintField;
                        result.Append(SqlNameMangling.GetSqlTableName(table.Name) + ".[" + SqlNameMangling.GetValidSqlNameForField(fieldConstraint.Field) + "]");
                        result.Append(" = ");
                        result.AppendLine(string.Format(CultureInfo.InvariantCulture, "t{0}", disambiguation) + "." + SqlNameMangling.GetValidSqlNameForField(fieldConstraint.RelatedField));
                    }
                    else if (constraint is TablesAutomation.IRelationConstraintFixed)
                    {   // Table.field = value
                        var fixedConstraint = constraint as TablesAutomation.IRelationConstraintFixed;

                        result.Append(SqlNameMangling.GetSqlTableName(table.Name) + ".[" + SqlNameMangling.GetValidSqlNameForField(fixedConstraint.Field) + "]");
                        result.Append(" = ");
                        result.AppendLine(fixedConstraint.Value.ToString(CultureInfo.InvariantCulture));
                    }
                    else if (constraint is TablesAutomation.IRelationConstraintRelatedFixed)
                    {   // Value = RelatedTable.field
                        var relatedFixedConstraint = constraint as TablesAutomation.IRelationConstraintRelatedFixed;
                        result.Append(relatedFixedConstraint.Value);
                        result.Append(" = ");
                        result.AppendLine(string.Format(CultureInfo.InvariantCulture, "t{0}", disambiguation) + ".[" + SqlNameMangling.GetValidSqlNameForField(relatedFixedConstraint.RelatedField) + "]");
                    }

                    first = false;
                }

                disambiguation += 1;
            }

            // If this table saves data per company, then add the where clause for
            // the user to fill out.
            if (tables.First().SaveDataPerCompany == Metadata.Core.MetaModel.NoYes.Yes)
            {
                // Note: There is no way to get to the actual company, since the server 
                // is not involved in this at all.
                result.AppendLine("-- where " + SqlNameMangling.GetValidSqlNameForField("DataAreaId") + " = 'DAT'");
            }

            return result;
        }

        #region Callbacks
        /// <summary>
        /// Called when user clicks on the add-in menu
        /// </summary>
        /// <param name="e">The context of the VS tools and metadata</param>
        public override void OnClick(AddinDesignerEventArgs e)
        {
            try
            {
                StringBuilder result = null;
                TablesAutomation.ITable selectedTable;

                if ((selectedTable = e.SelectedElement as TablesAutomation.ITable) != null)
                {
                    result = this.GenerateFromTable(selectedTable);
                }

                if (result == null)
                {
                    ViewsAutomation.IView selectedView;
                    if ((selectedView = e.SelectedElement as ViewsAutomation.IView) != null)
                    {
                        result = this.GenerateFromView(selectedView);
                    }
                }

                if (result == null)
                {
                    TablesAutomation.ITableExtension selectedTableExtension;
                    if ((selectedTableExtension = e.SelectedElement as TablesAutomation.ITableExtension) != null)
                    {
                        result = this.GenerateFromTableExtension(selectedTableExtension);
                    }
                }

                // Individually selected table fields.
                if (result == null)
                {
                    var selectedFields = e.SelectedElements.OfType<TablesAutomation.IBaseField>();
                    if (selectedFields.Any())
                    {
                        result = this.GenerateFromTableFieldListWithWhere(selectedFields);
                    }
                }

                // Individually selected view fields.
                if (result == null)
                {
                    var selectedFields = e.SelectedElements.OfType<ViewsAutomation.IViewBaseField>();
                    if (selectedFields.Any())
                    {
                        result = this.GenerateFromViewFieldList(selectedFields);
                    }
                }

                // Table field groups
                if (result == null)
                {
                    var selectedFieldsGroups = e.SelectedElements.OfType<TablesAutomation.IFieldGroup>();
                    if (selectedFieldsGroups.Any())
                    {
                        result = this.GenerateFromTableFieldGroups(selectedFieldsGroups);
                    }
                }

                // View field groups
                if (result == null)
                {
                    var selectedFieldsGroups = e.SelectedElements.OfType<ViewsAutomation.IFieldGroup>();
                    if (selectedFieldsGroups.Any())
                    {
                        result = this.GenerateFromViewFieldGroups(selectedFieldsGroups);
                    }
                }


                // Table relations
                if (result == null)
                {
                    var selectedRelations = e.SelectedElements.OfType<TablesAutomation.IRelation>();
                    if (selectedRelations.Any())
                    {
                        result = this.GenerateFromTableRelations(selectedRelations);
                    }
                }

                if (result != null)
                {
                    // Save the SQL file and open it in SQL management studio. 
                    string temporaryFileName = Path.GetTempFileName();

                    // Rename and move
                    var sqlFileName = temporaryFileName.Replace(".tmp", ".sql");
                    File.Move(temporaryFileName, sqlFileName);

                    // Store the script in the file
                    File.AppendAllText(sqlFileName, result.ToString());

                    //var dte = CoreUtility.ServiceProvider.GetService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                    //dte.ExecuteCommand("File.OpenFile", sqlFileName);

                    Process sqlManagementStudio = new Process();
                    sqlManagementStudio.StartInfo.FileName = sqlFileName;
                    sqlManagementStudio.StartInfo.UseShellExecute = true;
                    sqlManagementStudio.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                    sqlManagementStudio.Start();
                }
            }
            catch (Exception ex)
            {
                CoreUtility.HandleExceptionWithErrorMessage(ex);
            }
        }

        #endregion
    }
}
