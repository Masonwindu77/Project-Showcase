
protected void CustomRowCellStyle(object sender, ASPxGridViewTableDataCellEventArgs gridDataCellEventArg)
{           
	ASPxGridView gridViewSender = (sender as ASPxGridView);          
	CheckDataForCorrectFieldName(gridDataCellEventArg, gridViewSender);
}

protected void CheckDataForCorrectFieldName(ASPxGridViewTableDataCellEventArgs gridDataCellEventArg, ASPxGridView gridViewSender)
{
	if (gridDataCellEventArg.DataColumn.FieldName.Contains("OTD"))
	{
		SetOnTimeDeliveryFieldsCustomCellStyle(gridDataCellEventArg);
	}
	else if (gridDataCellEventArg.DataColumn.FieldName.Contains("DPPM"))
	{
		SetDefectPerPartMillionCustomCellStyle(gridDataCellEventArg);
	}
	else if (gridDataCellEventArg.DataColumn.FieldName.Contains("Defect"))
	{
		SetDefectsCustomCellStyle(gridDataCellEventArg);
	}
	else if (gridDataCellEventArg.DataColumn.FieldName.ToString() == "Vendor_Code")
	{
		SetPreferenceCustomCellStyle(gridDataCellEventArg, gridViewSender);
	}
	else if (gridDataCellEventArg.DataColumn.FieldName.ToString() == "Vendor_Name")
	{
		SetTSPCodeCustomCellStyle(gridDataCellEventArg, gridViewSender);
	}
}
