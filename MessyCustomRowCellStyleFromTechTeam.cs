/// <summary>
        /// Changes the color of cells depending on value
        /// Goes through the values of each cell and checks with the Metrics table on what the value should be
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="evCell"></param>
        protected void CustomRowCellStyle(object sender, ASPxGridViewTableDataCellEventArgs evCell)
        {
            string[] OTD = { "One_Month_OTD", "Three_Month_OTD", "Six_Month_OTD", "Twelve_Month_OTD" };
            string[] DPPM = { "One_Month_DPPM", "Three_Month_DPPM", "Six_Month_DPPM", "Twelve_Month_DPPM" };
            string[] Defect = { "One_Month_Defect", "Three_Month_Defect", "Six_Month_Defect", "Twelve_Month_Defect" };

            ASPxGridView gv = (sender as ASPxGridView);

            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = @"SELECT
                      OTD_High 
                     ,OTD_Low
                     ,DPPM_High
                     ,DPPM_Low
                     ,Defect_High 
                     ,Defect_Low
                     ,Supplier_Defect_High
                     ,Supplier_Defect_Low
                  FROM MS_SAP.TechTeam.Metrics
                     WHERE UserName = @userName 
                            OR UserName = 'Default'
                     ORDER BY ID DESC";
                cmd.Parameters.AddNullable("@userName", webUser.NetworkId.ToString().ToLower());
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    dt.Load(reader);
                }
                conn.Close();
            }

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < OTD.Count(); i++)
                {
                    decimal outValue;

                    if (evCell.DataColumn.FieldName == OTD[i].ToString() && !(dt.Rows[0]["OTD_High"].ToString() == "") && !(dt.Rows[0]["OTD_Low"].ToString() == ""))
                    {

                        if (decimal.TryParse(evCell.CellValue.ToString(), out outValue))
                        {
                            if (outValue >= (decimal.Parse(dt.Rows[0]["OTD_High"].ToString())) / 100)
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Green;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else if (outValue <= (decimal.Parse(dt.Rows[0]["OTD_Low"].ToString())) / 100)
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Red;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Yellow;
                            }
                        }

                    }
                    else if (evCell.DataColumn.FieldName == DPPM[i].ToString() && !(dt.Rows[0]["DPPM_High"].ToString() == "") && !(dt.Rows[0]["DPPM_Low"].ToString() == ""))
                    {

                        if (decimal.TryParse(evCell.CellValue.ToString(), out outValue))
                        {
                            if (outValue <= decimal.Parse(dt.Rows[0]["DPPM_Low"].ToString()))
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Green;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else if (outValue >= decimal.Parse(dt.Rows[0]["DPPM_High"].ToString()))
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Red;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Yellow;
                            }
                        }
                    }
                    else if (evCell.DataColumn.FieldName == Defect[i].ToString() && !(dt.Rows[0]["Defect_High"].ToString() == "") && !(dt.Rows[0]["Defect_Low"].ToString() == ""))
                    {
                        if (decimal.TryParse(evCell.CellValue.ToString(), out outValue))
                        {
                            if (outValue >= decimal.Parse(dt.Rows[0]["Defect_High"].ToString()))
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Red;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else if (outValue <= decimal.Parse(dt.Rows[0]["Defect_Low"].ToString()))
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Green;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Yellow;
                            }
                        }
                    }
                    else if (evCell.DataColumn.FieldName == "Supplier_Defect_Percent" && !(dt.Rows[0]["Supplier_Defect_High"].ToString() == "") && !(dt.Rows[0]["Supplier_Defect_Low"].ToString() == ""))
                    {

                        if (decimal.TryParse(evCell.CellValue.ToString(), out outValue))
                        {
                            if (outValue >= (decimal.Parse(dt.Rows[0]["Supplier_Defect_High"].ToString())) / 100)
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Red;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else if (outValue <= (decimal.Parse(dt.Rows[0]["Supplier_Defect_Low"].ToString())) / 100)
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Green;
                                evCell.Cell.ForeColor = System.Drawing.Color.White;
                            }
                            else
                            {
                                evCell.Cell.BackColor = System.Drawing.Color.Yellow;
                            }
                        }
                    }
                }
            }

            if (evCell.DataColumn.FieldName == "Vendor_Code")
            {
                if (gv.GetRowValues(evCell.VisibleIndex, "Preference") != null)
                {
                    if (gv.GetRowValues(evCell.VisibleIndex, "Preference").ToString() == "Preferred")
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Green;
                        evCell.Cell.ForeColor = System.Drawing.Color.White;
                    }
                    else if (gv.GetRowValues(evCell.VisibleIndex, "Preference").ToString() == "Non-Preferred")
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Yellow;
                    }
                    else if (gv.GetRowValues(evCell.VisibleIndex, "Preference").ToString() == "Exit")
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Red;
                        evCell.Cell.ForeColor = System.Drawing.Color.White;
                    }
                }
            }

            if (evCell.DataColumn.FieldName == "Vendor_Name")
            {
                if (gv.GetRowValues(evCell.VisibleIndex, "TSP_Code") != null)
                {
                    if (gv.GetRowValues(evCell.VisibleIndex, "TSP_Code").ToString().Contains("1"))
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Green;
                        evCell.Cell.ForeColor = System.Drawing.Color.White;
                    }
                    else if (gv.GetRowValues(evCell.VisibleIndex, "TSP_Code").ToString().Contains("2"))
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Gold;
                    }
                    else if (gv.GetRowValues(evCell.VisibleIndex, "TSP_Code").ToString().Contains("3"))
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Silver;
                        evCell.Cell.ForeColor = System.Drawing.Color.White;
                    }
                    else if (gv.GetRowValues(evCell.VisibleIndex, "TSP_Code").ToString().Contains("4"))
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Orange;
                        evCell.Cell.ForeColor = System.Drawing.Color.White;
                    }
                    else if (gv.GetRowValues(evCell.VisibleIndex, "TSP_Code").ToString().Contains("5"))
                    {
                        evCell.Cell.BackColor = System.Drawing.Color.Red;
                        evCell.Cell.ForeColor = System.Drawing.Color.White;
                    }
                }
            }
        }