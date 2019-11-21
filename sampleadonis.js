module.exports ={
	"model": {
		"groups": {
			"group": {
				"_id": "{3b7034d4-a336-4ed5-bef2-363d7a48c55e}",
				"__text": "/Modele/"
			}
		},
		"notebook": {
			"chapter": [
				{
					"attribute": [
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
									"_type": "ADOSTRING"
								},
								"value": {
									"p": "Przykładowy diagram IC (wersjonowanie)"
								},
								"_uuid": "{2d9b952b-d484-4a1a-86e6-5ad3fffab66a}"
							},
							"_id": "{f9a0c671-68d2-4703-bb93-aa466860b06d}",
							"_name": "Nazwa",
							"_idname": "NAME",
							"_type": "ADOSTRING",
							"_controltype": "NAME"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
									"_type": "BOOL"
								},
								"value": "0",
								"_uuid": "{3641f3ff-0262-4f54-9416-26e513b080a3}"
							},
							"_id": "{c1be28cd-6f70-46ec-aa34-b7bc576cc2fe}",
							"_name": "Proces wewnętrzny",
							"_idname": "A_INTERNAL_PROCESS",
							"_type": "BOOL",
							"_controltype": "BOOL"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
									"_type": "ENUM"
								},
								"_uuid": "{238527f6-ce9b-4244-a358-498869b14a35}",
								"_value": "v0",
								"_name": "Proces",
								"_constraint": "v0@v1",
								"_nameconstraint": "Proces@Podproces"
							},
							"_id": "{76f369f6-1011-45f2-8da9-7aec3ed22b8b}",
							"_name": "Typ procesu",
							"_idname": "A_PROCESS_TYPE_IC",
							"_type": "ENUM",
							"_controltype": "ENUM"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{bcfca9dc-c54f-4b73-9280-38b8ed8a7b10}",
									"_type": "LONGSTRING"
								},
								"value": "Tu jest podany cel procesu",
								"richvalue": {
									"x": "Tu jest podany cel procesu"
								},
								"_uuid": "{551638f2-b9d6-4521-a648-23fafabee14b}"
							},
							"_id": "{f060bf69-c844-4c8b-a217-4197bf68a3c5}",
							"_name": "Cel",
							"_idname": "AIM",
							"_type": "LONGSTRING",
							"_controltype": "STRING_MULTILINE"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{bcfca9dc-c54f-4b73-9280-38b8ed8a7b10}",
									"_type": "LONGSTRING"
								},
								"value": "Proces obejmuje następujący zakres...",
								"richvalue": {
									"x": "Proces obejmuje następujący zakres..."
								},
								"_uuid": "{e46876d8-a918-45ff-b3f9-c6169419105d}"
							},
							"_id": "{e9cd69f0-9e79-420a-84fe-e0b1d0bf40ad}",
							"_name": "Zakres procesu",
							"_idname": "A_PROCESS_SCOPE",
							"_type": "LONGSTRING",
							"_controltype": "STRING_MULTILINE"
						}
					],
					"_name": "Informacje ogólne"
				},
				{
					"relation": [
						{
							"link": {
								"endpoint": {
									"_type": "repoinst",
									"_role": "TO",
									"_objectid": "{bbfaff1d-1d53-4f0b-831b-825e54201c55}",
									"_name": "guest",
									"_class": "Użytkownik"
								},
								"_id": "{24baeb3b-df56-4be8-a149-5ad9644fa100}"
							},
							"_class": "Właściciel procesu",
							"_idclass": "RC_OWNER",
							"_label": "Właściciel procesu",
							"_controltype": "RELATIONS"
						},
						{
							"link": {
								"endpoint": {
									"_type": "repoinst",
									"_role": "TO",
									"_objectid": "{bbfaff1d-1d53-4f0b-831b-825e54201c55}",
									"_name": "guest",
									"_class": "Użytkownik"
								},
								"_id": "{637f4e78-061c-41ab-8a2d-28146d3fe4bf}"
							},
							"_class": "Lider procesu",
							"_idclass": "RC_MANAGER",
							"_label": "Lider procesu",
							"_controltype": "RELATIONS"
						},
						{
							"link": {
								"endpoint": {
									"_type": "repoinst",
									"_role": "TO",
									"_objectid": "{bbfaff1d-1d53-4f0b-831b-825e54201c55}",
									"_name": "guest",
									"_class": "Użytkownik"
								},
								"_id": "{dd604be6-b0e4-4414-a677-080f4230cba3}"
							},
							"_class": "Recenzja metodyczna",
							"_idclass": "RC_METHODICAL_REVIEWER",
							"_label": "Recenzja metodyczna",
							"_controltype": "RELATIONS"
						},
						{
							"link": {
								"endpoint": {
									"_type": "repoinst",
									"_role": "TO",
									"_objectid": "{bbfaff1d-1d53-4f0b-831b-825e54201c55}",
									"_name": "guest",
									"_class": "Użytkownik"
								},
								"_id": "{941d757f-122f-419f-966a-6fc1d9f6bdff}"
							},
							"_class": "Analityk procesu",
							"_idclass": "RC_ANALYST",
							"_label": "Analityk procesu",
							"_controltype": "RELATIONS"
						}
					],
					"_name": "Organizacja"
				},
				{
					"_name": "Klasyfikacja"
				},
				{
					"group": {
						"_name": "Ocena stanu obecnego"
					},
					"_name": "Analiza potencjału i dojrzałości"
				},
				{
					"relation": [
						{
							"link": {
								"endpoint": {
									"_type": "repoinst",
									"_role": "TO",
									"_objectid": "{c0e6cee6-3ad7-4c84-9fe9-f7573fbeafe0}",
									"_name": "Dokument A ",
									"_class": "Dokument"
								},
								"_id": "{d7b2fc22-561d-4146-983d-28372ea9288f}"
							},
							"_class": "Wejście",
							"_idclass": "RC_INPUT_MT",
							"_label": "Wejście",
							"_controltype": "RELATIONS"
						},
						{
							"link": {
								"endpoint": {
									"_type": "repoinst",
									"_role": "TO",
									"_objectid": "{ce0dd4e5-8607-44bb-a391-5966b8a088f1}",
									"_name": "Dokument B ",
									"_class": "Dokument"
								},
								"_id": "{e88f6120-f563-4bfc-9698-2a05b83d6740}"
							},
							"_class": "Wyjście",
							"_idclass": "RC_OUTPUT_MT",
							"_label": "Wyjście",
							"_controltype": "RELATIONS"
						}
					],
					"_name": "Wejście/wyjście"
				},
				{
					"relation": {
						"link": {
							"endpoint": {
								"_type": "repoinst",
								"_role": "TO",
								"_objectid": "{d353e2c5-09ac-4ee9-b59e-510a30216f50}",
								"_name": "Lista kontrolna ABC ",
								"_class": "Dokument"
							},
							"_id": "{291295ff-1a31-4aaa-8f89-76d82c78e990}"
						},
						"_class": "Powiązane dokumenty",
						"_idclass": "RC_REFERENCED_DOCUMENTS_MT",
						"_label": "Powiązane dokumenty",
						"_controltype": "RELATIONS"
					},
					"group": {
						"_name": "Terminy, skróty"
					},
					"_name": "Dokumenty"
				},
				{
					"attribute": [
						{
							"columns": {
								"member": [
									{
										"_name": "Jednostka organizacyjna/Partner zewnętrzny"
									},
									{
										"_name": "Wymagania"
									}
								]
							},
							"complexvalues": {
								"member": {
									"attrval": {
										"attrvaltype": {
											"_uuid": "{d3429e04-b06e-49be-9fb4-e7252ecff8d6}"
										},
										"_uuid": "{faf84880-dcc6-4023-85f5-909e61a5130e}"
									},
									"complexvalues": {
										"member": [
											{
												"attrval": {
													"attrvaltype": {
														"_uuid": "{b570c563-6c13-4f2f-ae68-1a5b8ca123f3}",
														"_type": "INTERREF"
													},
													"relation": {
														"link": {
															"endpoint": {
																"_type": "repoinst",
																"_role": "TO",
																"_objectid": "{67f0fd50-90eb-4f5d-92e1-577fffeda9e9}",
																"_name": "Rola A",
																"_class": "Rola"
															},
															"_id": "{d74bdb0c-bab7-48b3-a78b-68ffae510c66}"
														},
														"_class": "Jednostka organizacyjna/Partner zewnętrzny",
														"_idclass": "ORGANISATION_UNITEXTERNAL_PARTNER"
													},
													"_uuid": "{b56c0911-7f41-48f8-8b26-3ef7d703be1d}"
												},
												"_complex": "0",
												"_name": "Jednostka organizacyjna/Partner zewnętrzny"
											},
											{
												"attrval": {
													"attrvaltype": {
														"_uuid": "{eb509df5-0d10-42ab-b6bf-19d3110b34df}",
														"_type": "STRING"
													},
													"value": {
														"p": "Wymaganie A"
													},
													"_uuid": "{ef9408eb-c9cb-401d-921a-19004995cb15}"
												},
												"_complex": "0",
												"_name": "Wymagania"
											}
										]
									},
									"_complex": "1",
									"_name": "Columns"
								}
							},
							"_id": "{cf2c8992-931e-46f3-bcdc-ee976e9eddd3}",
							"_name": "Klienci procesu - wymagania",
							"_idname": "PROCESS_CUSTOMERS___REQUIREMENTS",
							"_type": "CUSTOMERSUPPLIER_REQUIREMENTS",
							"_complex": "1",
							"_controltype": "RECORD"
						},
						{
							"columns": {
								"member": [
									{
										"_name": "Jednostka organizacyjna/Partner zewnętrzny"
									},
									{
										"_name": "Wymagania"
									}
								]
							},
							"complexvalues": {
								"member": [
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{d3429e04-b06e-49be-9fb4-e7252ecff8d6}"
											},
											"_uuid": "{3e6f48fc-5d60-4a89-9c73-b7eefa1b1a7e}"
										},
										"complexvalues": {
											"member": [
												{
													"attrval": {
														"attrvaltype": {
															"_uuid": "{b570c563-6c13-4f2f-ae68-1a5b8ca123f3}",
															"_type": "INTERREF"
														},
														"relation": {
															"link": {
																"endpoint": {
																	"_type": "repoinst",
																	"_role": "TO",
																	"_objectid": "{a43c3c3f-f3a4-4205-97ca-7ab9015765c6}",
																	"_name": "Rola B",
																	"_class": "Rola"
																},
																"_id": "{01b8f97f-a4f9-4d5d-92bf-ec66471d85a4}"
															},
															"_class": "Jednostka organizacyjna/Partner zewnętrzny",
															"_idclass": "ORGANISATION_UNITEXTERNAL_PARTNER"
														},
														"_uuid": "{4c8ef8fa-aef2-4b87-ac2f-5bf7b578e7e9}"
													},
													"_complex": "0",
													"_name": "Jednostka organizacyjna/Partner zewnętrzny"
												},
												{
													"attrval": {
														"attrvaltype": {
															"_uuid": "{eb509df5-0d10-42ab-b6bf-19d3110b34df}",
															"_type": "STRING"
														},
														"value": {
															"p": "Wymaganie B"
														},
														"_uuid": "{66094c6d-eb5f-4629-9888-b1fc03612621}"
													},
													"_complex": "0",
													"_name": "Wymagania"
												}
											]
										},
										"_complex": "1",
										"_name": "Columns"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{d3429e04-b06e-49be-9fb4-e7252ecff8d6}"
											},
											"_uuid": "{03b3bc81-b405-4a47-8291-04bcdbf6060c}"
										},
										"complexvalues": {
											"member": [
												{
													"attrval": {
														"attrvaltype": {
															"_uuid": "{b570c563-6c13-4f2f-ae68-1a5b8ca123f3}",
															"_type": "INTERREF"
														},
														"relation": {
															"link": {
																"endpoint": {
																	"_type": "repoinst",
																	"_role": "TO",
																	"_objectid": "{f97cab4c-95a7-4b65-9dd3-7322898bae86}",
																	"_name": "Rola C",
																	"_class": "Rola"
																},
																"_id": "{8f065ed7-2d8c-4049-a7e0-37287ecc4ccc}"
															},
															"_class": "Jednostka organizacyjna/Partner zewnętrzny",
															"_idclass": "ORGANISATION_UNITEXTERNAL_PARTNER"
														},
														"_uuid": "{1a2154d6-f10c-4a7e-b571-0bf6031ac069}"
													},
													"_complex": "0",
													"_name": "Jednostka organizacyjna/Partner zewnętrzny"
												},
												{
													"attrval": {
														"attrvaltype": {
															"_uuid": "{eb509df5-0d10-42ab-b6bf-19d3110b34df}",
															"_type": "STRING"
														},
														"value": {
															"p": "Wymaganie C"
														},
														"_uuid": "{87aba090-aae7-4595-8b08-dd3ca258097f}"
													},
													"_complex": "0",
													"_name": "Wymagania"
												}
											]
										},
										"_complex": "1",
										"_name": "Columns"
									}
								]
							},
							"_id": "{8a716e31-e16f-4192-8dfa-1568d51babdc}",
							"_name": "Dostawcy procesu - wymagania",
							"_idname": "PROCESS_SUPPLIERS___REQUIREMENTS",
							"_type": "CUSTOMERSUPPLIER_REQUIREMENTS",
							"_complex": "1",
							"_controltype": "RECORD"
						}
					],
					"_name": "Klienci/dostawcy"
				},
				{
					"group": {
						"_name": "Wymagana akcja"
					},
					"_name": "Ciągłe usprawnianie"
				},
				{
					"relation": {
						"link": {
							"endpoint": {
								"_type": "repoinst",
								"_role": "TO",
								"_objectid": "{8d39c5d8-f047-4619-a097-d0c2184cc3c7}",
								"_name": "Ryzyko A",
								"_class": "Ryzyko"
							},
							"_id": "{87cc19cb-3305-4203-901e-e6266ebfcfff}"
						},
						"_class": "Ryzyka",
						"_idclass": "RC_ASSIGNED_RISKS_MT",
						"_label": "Ryzyka",
						"_controltype": "RELATIONS"
					},
					"_name": "Zarządzanie ryzykiem"
				},
				{
					"_name": "Zgodność"
				},
				{
					"relation": {
						"link": {
							"endpoint": {
								"_type": "repoinst",
								"_role": "TO",
								"_objectid": "{218f4157-c6da-43d0-b330-2bab8a8c42b0}",
								"_name": "Produkt A",
								"_class": "Produkt"
							},
							"_id": "{8a17eda6-b526-469d-a163-151ee5302218}"
						},
						"_class": "Produkty",
						"_idclass": "RC_REFERENCED_PRODUCTS_MT",
						"_label": "Produkty",
						"_controltype": "RELATIONS"
					},
					"_name": "Systemy/produkty"
				},
				{
					"group": [
						{
							"_name": "Atrybuty procesu"
						},
						{
							"_name": "Znaczenie dla procesu"
						},
						{
							"_name": "Atrybuty choreografii"
						}
					],
					"_name": "Właściwości"
				},
				{
					"attribute": [
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
									"_type": "ENUM"
								},
								"_uuid": "{520e507c-c69a-4b84-b3ea-dc3e1245e58e}",
								"_value": "v0",
								"_name": "szkic",
								"_constraint": "v0@v7@v8@v12@v9@v10@v11@v3@v2@v4",
								"_nameconstraint": "szkic@założony@w trakcie analizy@do aktualizacji@do publikacji@zatwierdzony@zatwierdzony do zarządzania jakością@archiwalny@aktualny@nieaktualny"
							},
							"_id": "{7dc7da32-56a4-426c-bf0c-d9020e6b277e}",
							"_name": "Status",
							"_idname": "MFB_RWF_STATE",
							"_type": "ENUM",
							"_controltype": "ENUM"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{c6b095c1-f03e-4268-b0d2-4e60c1c2e78e}",
									"_type": "STRING"
								},
								"value": {
									"p": "0.01"
								},
								"_uuid": "{4f0e4f59-4dcd-4533-bd0e-867f4b5e234b}"
							},
							"_id": "{d8d180d5-8329-4c3a-9d19-0d67055a9b93}",
							"_name": "Wersja",
							"_idname": "MFB_RWF_VERSION",
							"_type": "STRING",
							"_controltype": "STRING"
						},
						{
							"columns": {
								"member": [
									{
										"_name": "Komentarz"
									},
									{
										"_name": "Data"
									},
									{
										"_name": "Użytkownik"
									},
									{
										"_name": "Wersja modelu"
									},
									{
										"_name": "Status modelu"
									}
								]
							},
							"complexvalues": {
								"member": {
									"attrval": {
										"attrvaltype": {
											"_uuid": "{fbec863c-eab4-418f-a81d-d73f796ea16a}"
										},
										"_uuid": "{8cc03470-4b10-49c4-92a4-71331ed2447a}"
									},
									"complexvalues": {
										"member": [
											{
												"attrval": {
													"attrvaltype": {
														"_uuid": "{667b19e2-f759-4b37-9e6d-a3b3d78e8ab6}",
														"_type": "LONGSTRING"
													},
													"value": {
														"p": "Utworzono nowy model."
													},
													"_uuid": "{20471f63-8f02-481e-89b2-c1912c191686}"
												},
												"_complex": "0",
												"_name": "Komentarz"
											},
											{
												"attrval": {
													"attrvaltype": {
														"_uuid": "{3b461abd-85e9-49fc-b173-4b924e4e88c9}",
														"_type": "UTC"
													},
													"value": "1564481920490",
													"alternate-value": "30.07.2019 12:18",
													"_uuid": "{76d44a2e-a899-4b29-848c-e043807702eb}"
												},
												"_complex": "0",
												"_name": "Data"
											},
											{
												"attrval": {
													"attrvaltype": {
														"_uuid": "{fbe6e369-25ca-4fdc-830a-89cb51a458ea}",
														"_type": "STRING"
													},
													"value": {
														"p": "guest"
													},
													"_uuid": "{ab6316a3-7ac3-4d71-a85a-2aede9650e8a}"
												},
												"_complex": "0",
												"_name": "Użytkownik"
											},
											{
												"attrval": {
													"attrvaltype": {
														"_uuid": "{2512cc79-7dbc-42c6-ab14-2f4eee40e977}",
														"_type": "STRING"
													},
													"value": {
														"p": "0.01"
													},
													"_uuid": "{c8eac3a8-d0c3-4697-b9f5-7aae57f8d69a}"
												},
												"_complex": "0",
												"_name": "Wersja modelu"
											},
											{
												"attrval": {
													"attrvaltype": {
														"_uuid": "{67e1771f-d294-4f3f-9b8d-643b0997ef06}",
														"_type": "ENUM"
													},
													"_uuid": "{7a2cf2cd-4a51-4674-8547-f1653f525305}",
													"_value": "v0",
													"_name": "szkic",
													"_constraint": "v0@v1@v6@v5@v2@v4@v3",
													"_nameconstraint": "szkic@w trakcie recenzji metodycznej@w trakcie recenzji biznesowej@zaakceptowany@aktualny@nieaktualny@archiwalny"
												},
												"_complex": "0",
												"_name": "Status modelu"
											}
										]
									},
									"_complex": "1",
									"_name": "Columns"
								}
							},
							"_id": "{0e7effb3-b1c6-418d-963c-53a8941f3386}",
							"_name": "Historia wersji",
							"_idname": "MFB_RWF_HISTORY",
							"_type": "MFB_RWF_HISTORY_AT",
							"_complex": "1",
							"_controltype": "RECORD"
						}
					],
					"_name": "Cykl życia"
				},
				{
					"attribute": {
						"columns": {
							"member": [
								{
									"_name": "Link"
								},
								{
									"_name": "Komentarz"
								}
							]
						},
						"complexvalues": {
							"member": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{c22ceb51-5f77-47c5-8a84-dafb7b545f6d}"
									},
									"_uuid": "{07f918fa-a8cd-4ccf-b64b-7950293415fe}"
								},
								"complexvalues": {
									"member": [
										{
											"attrval": {
												"attrvaltype": {
													"_uuid": "{03360474-963a-424a-893c-832687c75e8f}",
													"_type": "FILE_POINTER"
												},
												"value": "\"\" \"https://intercars.com.pl/\"",
												"_uuid": "{8a82fbda-6263-4218-ae55-c70970cd915f}"
											},
											"_complex": "0",
											"_name": "Link"
										},
										{
											"attrval": {
												"attrvaltype": {
													"_uuid": "{a5c5e755-2dbd-44a5-8c50-f8225ab81795}",
													"_type": "STRING"
												},
												"value": {
													"p": "Strona WWW"
												},
												"_uuid": "{81d9b70a-06c7-462c-9e7a-39e83cd6ae65}"
											},
											"_complex": "0",
											"_name": "Komentarz"
										}
									]
								},
								"_complex": "1",
								"_name": "Columns"
							}
						},
						"_id": "{31924820-f1c9-4634-941b-0a25aff4456f}",
						"_name": "Linki",
						"_idname": "A_ATTACHMENT_LIST",
						"_type": "FILE_POINTER_LIST",
						"_complex": "1",
						"_controltype": "RECORD"
					},
					"_name": "Załączniki"
				},
				{
					"_name": "Wykonanie procesu"
				},
				{
					"attribute": [
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
									"_type": "ADOSTRING"
								},
								"value": "guest",
								"richvalue": "guest",
								"_uuid": "{b0e0cdde-9401-4fb7-a03c-016484f3af16}"
							},
							"_id": "{b099665d-b0ae-4b3f-bd89-07a35f5b4b1d}",
							"_name": "Autor",
							"_idname": "AUTHOR",
							"_type": "ADOSTRING",
							"_controltype": "STRING"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{3b89a252-3f2e-4559-884f-1c7496a6cc21}",
									"_type": "DOUBLE"
								},
								"value": "1564479683701",
								"alternate-value": "30.07.2019 11:41",
								"_uuid": "{b73e29a6-a0da-40d9-8d2c-e1b84334d477}"
							},
							"_id": "{ed3fde02-ddd6-4685-bd88-930492412d70}",
							"_name": "Data utworzenia",
							"_idname": "CREATION_DATE",
							"_type": "DOUBLE",
							"_controltype": "UTC"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
									"_type": "ADOSTRING"
								},
								"value": "guest",
								"richvalue": "guest",
								"_uuid": "{e196e5d3-1b12-491a-bb8f-6550b51b665e}"
							},
							"_id": "{eafafc5f-cb8a-47f4-866a-b20d219d814c}",
							"_name": "Autor ostatniej zmiany",
							"_idname": "LAST_EDITOR",
							"_type": "ADOSTRING",
							"_controltype": "STRING"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{3b89a252-3f2e-4559-884f-1c7496a6cc21}",
									"_type": "DOUBLE"
								},
								"value": "1564481965873",
								"alternate-value": "30.07.2019 12:19",
								"_uuid": "{03672af5-7628-4888-b2f3-a19f42b81aa8}"
							},
							"_id": "{2867c669-8924-410a-82ba-17b35aa5d90d}",
							"_name": "Data ostatniej zmiany",
							"_idname": "DATE_OF_LAST_CHANGE",
							"_type": "DOUBLE",
							"_controltype": "UTC"
						},
						{
							"attrval": {
								"attrvaltype": {
									"_uuid": "{d5dac928-b19c-484e-b927-c47206df0911}",
									"_type": "UNSIGNED INTEGER"
								},
								"value": "19",
								"_uuid": "{df48e710-9291-4e34-9a2c-221595120940}"
							},
							"_id": "{5ccb0f26-2a84-4182-9266-70a4042a65fc}",
							"_name": "Liczba obiektów i relacji",
							"_idname": "NUMBER_OF_OBJECTS_AND_RELATIONS",
							"_type": "UNSIGNED INTEGER",
							"_controltype": "INTEGER"
						}
					],
					"_name": "Informacje systemowe"
				}
			]
		},
		"object": [
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zdarzenie końcowe"
										},
										"_uuid": "{8f8d4c21-4fe0-4fa5-b594-df7cdfc5f544}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "09",
										"_uuid": "{9d5b5765-681c-4c5c-94c6-93e27a882e08}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"group": {
								"_name": "Wyzwalacz"
							},
							"_name": "Typ zdarzenia"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Szczegóły importu"
								},
								{
									"_name": "Wejście danych"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"group": [
								{
									"_name": "Komunikat (szczegóły)"
								},
								{
									"_name": "Sygnał (szczegóły)"
								},
								{
									"_name": "Eskalacja (szczegóły)"
								},
								{
									"_name": "Błąd (szczegóły)"
								},
								{
									"_name": "Kompensacja (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"group": {
								"_name": "Wymgana akcja"
							},
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
										"_type": "BOOL"
									},
									"value": "1",
									"_uuid": "{92c726c4-70c2-4281-bc7c-68ed25f84042}"
								},
								"_id": "{e9ce8c1f-1e31-454a-b3a0-701b490c92ec}",
								"_name": "Pokaż nazwę",
								"_idname": "DISPLAY_NAME",
								"_type": "BOOL",
								"_controltype": "BOOL"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{129a3e24-1916-4199-87f0-82193de5d4c4}",
				"_name": "Zdarzenie końcowe",
				"_class": "Zdarzenie końcowe",
				"_idclass": "C_END_EVENT",
				"_repoid": "{0de3a47d-5223-4294-9155-77876498fb88}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Bramka niewykluczająca"
										},
										"_uuid": "{c4f83889-f3b7-4f08-a1a1-16fd79093f75}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "05",
										"_uuid": "{37b36297-9424-4281-9ddb-8f4a55f67e1a}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
											"_type": "ENUM"
										},
										"_uuid": "{5b3c216c-555f-4a21-a898-5a7dfc78745d}",
										"_value": "v1",
										"_name": "równoległa",
										"_constraint": "v0@v1@v2",
										"_nameconstraint": "alternatywna@równoległa@złożona"
									},
									"_id": "{005753e1-88fc-4065-ab1a-516b3f43e1b2}",
									"_name": "Typ bramki",
									"_idname": "TYPE_NON_EXCLUSIVE_GATEWAY",
									"_type": "ENUM",
									"_controltype": "ENUM"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
											"_type": "ENUM"
										},
										"_uuid": "{dff39643-79ed-46e0-abdb-aa9b3f465bc7}",
										"_value": "v0",
										"_name": "sterowana danymi",
										"_constraint": "v0@v1",
										"_nameconstraint": "sterowana danymi@sterowana zdarzeniami (tworzy instancje)"
									},
									"_id": "{17c6681b-9ae8-4b7c-a44d-00b4deef95e9}",
									"_name": "Typ (bramka równoległa)",
									"_idname": "TYPE_NON_EXCLUSIVE_GATEWAY_PARALLEL",
									"_type": "ENUM",
									"_controltype": "ENUM"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{f12669f8-b9ec-458d-a8a2-a370352bec8d}",
									"_value": "v2",
									"_name": "nie pokazuj",
									"_constraint": "v0@v1@v2",
									"_nameconstraint": "nad@pod@nie pokazuj"
								},
								"_id": "{2c4f5971-60c8-4c45-b347-736c50f35f60}",
								"_name": "Pokaż nazwę",
								"_idname": "REPRESENTATION_NAME_GATEWAY",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{23dd883a-447e-4c73-ba6b-bb66b09507de}",
				"_name": "Bramka niewykluczająca",
				"_class": "Bramka niewykluczająca",
				"_idclass": "C_NON_EXCLUSIVE_GATEWAY",
				"_repoid": "{ebdb3f64-4c1a-4401-b9c6-7aecc89ba3b0}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zdarzenie początkowe"
										},
										"_uuid": "{61709c2e-2d7a-4e5a-8967-e87e1d5c0e5b}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "01",
										"_uuid": "{c84e37ea-d40e-4edb-8cb2-b95ef642c07b}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{6bc7fd41-3835-40cd-9315-bcae9e7ed2ed}",
									"_value": "v0",
									"_name": "główny poziom",
									"_constraint": "v0@v1@v2",
									"_nameconstraint": "główny poziom@Pod-Proces wywoływany zdarzeniem (nie przerywający)@Pod-Proces wywoływany zdarzeniem (przerywający)"
								},
								"_id": "{232f88a5-fc3a-4a37-a75f-f7f60136b42c}",
								"_name": "Typ",
								"_idname": "EVENT_TYPE",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"group": [
								{
									"_name": "Wyzwalacz"
								},
								{
									"_name": "Wyzwalacz złożony"
								}
							],
							"_name": "Typ zdarzenia"
						},
						{
							"group": {
								"_name": "Wyjście (w trakcie wykonywania)"
							},
							"_name": "Wejście/wyjście"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Ciągłe usprawnianie"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Szczegóły importu"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"group": [
								{
									"_name": "Komunikat (szczegóły)"
								},
								{
									"_name": "Czas (szczegóły)"
								},
								{
									"_name": "Warunek (szczegóły)"
								},
								{
									"_name": "Sygnał (szczegóły)"
								},
								{
									"_name": "Eskalacja (szczegóły)"
								},
								{
									"_name": "Błąd (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"_name": "Dane symulacyjne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
										"_type": "BOOL"
									},
									"value": "1",
									"_uuid": "{c8f02836-db88-4448-b3c5-268581505d89}"
								},
								"_id": "{e9ce8c1f-1e31-454a-b3a0-701b490c92ec}",
								"_name": "Pokaż nazwę",
								"_idname": "DISPLAY_NAME",
								"_type": "BOOL",
								"_controltype": "BOOL"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{33022123-4660-4893-bb5f-0d7a6d9b2371}",
				"_name": "Zdarzenie początkowe",
				"_class": "Zdarzenie początkowe",
				"_idclass": "C_START_EVENT",
				"_repoid": "{374596cd-18db-44a1-867e-fec933a3ba38}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zadanie F"
										},
										"_uuid": "{1201950d-083f-4e17-a7b3-851e1de4fe90}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "08",
										"_uuid": "{0bc34a3d-9d94-4a1d-ab5b-9eccceb62dda}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"_name": "Diagram odpowiedzialności (RACI)"
						},
						{
							"_name": "Zasoby"
						},
						{
							"group": {
								"_name": "Wejście/Wyjście (w trakcie wykonywania)"
							},
							"_name": "Wejście/Wyjście"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Ciągłe usprawnianie"
						},
						{
							"_name": "Zarządzanie ryzykiem"
						},
						{
							"_name": "Kontrole"
						},
						{
							"_name": "Systemy/Produkty"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Powiązane obiekty"
								},
								{
									"_name": "Definicja obiektu"
								},
								{
									"_name": "Szczegóły importu"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"_name": "Pętla - standardowa"
						},
						{
							"group": {
								"_name": "Wywoływanie zdarzeń"
							},
							"_name": "Pętla - wiele instancji"
						},
						{
							"group": [
								{
									"_name": "Szczegóły implementacji"
								},
								{
									"_name": "Usługa (szczegóły)"
								},
								{
									"_name": "Wysłanie/Odbiór (szczegóły)"
								},
								{
									"_name": "Użytkownika (szczegóły)"
								},
								{
									"_name": "Skrypt (szczegóły)"
								},
								{
									"_name": "Wywołanie aktywności (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"_name": "Dane symulacyjne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{6e1ca5d9-2b44-435e-90cc-75a69c799157}",
									"_value": "v0",
									"_name": "w środku",
									"_constraint": "v0@v1",
									"_nameconstraint": "w środku@poniżej"
								},
								"_id": "{565e3a93-18d2-4518-bbb3-2dcd433e85d0}",
								"_name": "Pokaż nazwę",
								"_idname": "REPRESENTATION_NAME_TASK",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"group": {
								"attribute": [
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{4caf788d-c93c-48b8-a8b8-c7510cbe7718}"
										},
										"_id": "{7abda636-e481-4d1f-8a8f-debaa7ed8394}",
										"_name": "Pokaż role odpowiedzialne za wykonanie",
										"_idname": "DISPLAY_RESPONSIBLE_FOR_EXECUTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{d38bab97-c4e8-4b1b-9047-3bf98b0f8067}"
										},
										"_id": "{22c826d9-af67-49a0-8cc5-c536b4e9caa9}",
										"_name": "Pokaż symbol (\"i\"), jeżeli istnieje opis",
										"_idname": "DISPLAY_SYMBOL_IF_DESCRIPTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{15c9ff35-f72f-4e70-a766-b7597d2fa279}"
										},
										"_id": "{0c4d83d5-56f7-4eed-b897-6368f9067750}",
										"_name": "Pokaż symbol gdy ryzyka istnieją",
										"_idname": "DISPLAY_SYMBOLS_IF_RISKS_OR_CONTROLS",
										"_type": "BOOL",
										"_controltype": "BOOL"
									}
								],
								"_name": "Wizualizacja atrybutów"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{39cc6abd-eab9-466c-8b32-1536ff901a0a}",
				"_name": "Zadanie F",
				"_class": "Zadanie",
				"_idclass": "C_TASK",
				"_repoid": "{04accc91-0388-4fbf-bef5-a250d9bbb023}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zadanie B"
										},
										"_uuid": "{1ba4b679-a3e1-46e8-bc92-77bec795bf2f}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "03",
										"_uuid": "{769f1b3b-23f7-4e85-9e27-a5ef7a72f15d}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{bcfca9dc-c54f-4b73-9280-38b8ed8a7b10}",
											"_type": "LONGSTRING"
										},
										"value": "Opis B\nDruga linia\nI ostatnia",
										"richvalue": {
											"x": [
												"Opis B",
												"Druga linia",
												{
													"span": {
														"em": {
															"strong": "ostatnia"
														},
														"_style": "color: #ff0000;"
													},
													"__text": "I"
												}
											],
											"br": [
												"",
												""
											]
										},
										"_uuid": "{273d9eef-4832-4b9c-b45b-9370c120c306}"
									},
									"_id": "{2fcdce9b-4a60-4011-a3fb-b435ba8d2a6a}",
									"_name": "Opis",
									"_idname": "A_DESCRIPTION_CS",
									"_type": "LONGSTRING",
									"_controltype": "STRING_MULTILINE"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"relation": {
								"link": {
									"endpoint": {
										"_type": "repoinst",
										"_role": "TO",
										"_objectid": "{a43c3c3f-f3a4-4205-97ca-7ab9015765c6}",
										"_name": "Rola B",
										"_class": "Rola"
									},
									"_id": "{d5730367-7177-420c-9271-bb70fa715e50}"
								},
								"_class": "Odpowiedzialny za wykonanie",
								"_idclass": "RESPONSIBLE_FOR_EXECUTION",
								"_label": "Odpowiedzialny za wykonanie",
								"_controltype": "RELATIONS"
							},
							"_name": "Diagram odpowiedzialności (RACI)"
						},
						{
							"relation": [
								{
									"link": {
										"endpoint": {
											"_type": "repoinst",
											"_role": "TO",
											"_objectid": "{ce0dd4e5-8607-44bb-a391-5966b8a088f1}",
											"_name": "Dokument B ",
											"_class": "Dokument"
										},
										"_id": "{b4d983a5-3887-4518-8788-5ccc99237696}"
									},
									"_class": "Powiązane dokumenty",
									"_idclass": "RC_REFERENCED_DOCUMENTS_M",
									"_label": "Powiązane dokumenty",
									"_controltype": "RELATIONS"
								},
								{
									"link": [
										{
											"endpoint": {
												"_type": "repoinst",
												"_role": "TO",
												"_objectid": "{0b60b209-391a-4e7d-8f92-373ebc8c5b9a}",
												"_name": "Aplikacja A",
												"_class": "Aplikacja"
											},
											"_id": "{0c8b773d-72a6-410f-a694-544d72497788}"
										},
										{
											"endpoint": {
												"_type": "repoinst",
												"_role": "TO",
												"_objectid": "{a5a4f90d-156d-4d5f-acaf-c803cb7fdd27}",
												"_name": "Aplikacja B",
												"_class": "Aplikacja"
											},
											"_id": "{4d398f8f-4fde-42ca-ac44-7f2df54a6ae6}"
										}
									],
									"_class": "Wsparcie IT",
									"_idclass": "RC_REFERENCED_IT_SYSTEM_ELEMENTS_CS",
									"_label": "Wsparcie IT",
									"_controltype": "RELATIONS"
								}
							],
							"_name": "Zasoby"
						},
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{bcfca9dc-c54f-4b73-9280-38b8ed8a7b10}",
											"_type": "LONGSTRING"
										},
										"value": "Wymaganie jakościowe a",
										"richvalue": {
											"x": "Wymaganie jakościowe a"
										},
										"_uuid": "{dc6e423f-0016-4af4-b850-78b8305a05e2}"
									},
									"_id": "{598a534c-876b-496e-ba22-3cf45cbf4eea}",
									"_name": "Wejście - wymagania jakościowe",
									"_idname": "INPUT_QUALITY_REQUIREMENTS",
									"_type": "LONGSTRING",
									"_controltype": "STRING_MULTILINE"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{bcfca9dc-c54f-4b73-9280-38b8ed8a7b10}",
											"_type": "LONGSTRING"
										},
										"value": "Wymaganie b",
										"richvalue": {
											"x": "Wymaganie b"
										},
										"_uuid": "{ee5d56a3-0aac-4aa7-8886-6b65ed69d66b}"
									},
									"_id": "{f72b7e22-bad0-4258-a424-352fbe6ce5df}",
									"_name": "Wyjście - wymagania jakościowe",
									"_idname": "OUTPUT_QUALITY_REQUIREMENTS",
									"_type": "LONGSTRING",
									"_controltype": "STRING_MULTILINE"
								}
							],
							"group": {
								"_name": "Wejście/Wyjście (w trakcie wykonywania)"
							},
							"_name": "Wejście/Wyjście"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Ciągłe usprawnianie"
						},
						{
							"_name": "Zarządzanie ryzykiem"
						},
						{
							"_name": "Kontrole"
						},
						{
							"_name": "Systemy/Produkty"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Powiązane obiekty"
								},
								{
									"_name": "Definicja obiektu"
								},
								{
									"_name": "Szczegóły importu"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"_name": "Pętla - standardowa"
						},
						{
							"group": {
								"_name": "Wywoływanie zdarzeń"
							},
							"_name": "Pętla - wiele instancji"
						},
						{
							"group": [
								{
									"_name": "Szczegóły implementacji"
								},
								{
									"_name": "Usługa (szczegóły)"
								},
								{
									"_name": "Wysłanie/Odbiór (szczegóły)"
								},
								{
									"_name": "Użytkownika (szczegóły)"
								},
								{
									"_name": "Skrypt (szczegóły)"
								},
								{
									"_name": "Wywołanie aktywności (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"_name": "Dane symulacyjne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{6e1ca5d9-2b44-435e-90cc-75a69c799157}",
									"_value": "v0",
									"_name": "w środku",
									"_constraint": "v0@v1",
									"_nameconstraint": "w środku@poniżej"
								},
								"_id": "{565e3a93-18d2-4518-bbb3-2dcd433e85d0}",
								"_name": "Pokaż nazwę",
								"_idname": "REPRESENTATION_NAME_TASK",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"group": {
								"attribute": [
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{4caf788d-c93c-48b8-a8b8-c7510cbe7718}"
										},
										"_id": "{7abda636-e481-4d1f-8a8f-debaa7ed8394}",
										"_name": "Pokaż role odpowiedzialne za wykonanie",
										"_idname": "DISPLAY_RESPONSIBLE_FOR_EXECUTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{d38bab97-c4e8-4b1b-9047-3bf98b0f8067}"
										},
										"_id": "{22c826d9-af67-49a0-8cc5-c536b4e9caa9}",
										"_name": "Pokaż symbol (\"i\"), jeżeli istnieje opis",
										"_idname": "DISPLAY_SYMBOL_IF_DESCRIPTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{15c9ff35-f72f-4e70-a766-b7597d2fa279}"
										},
										"_id": "{0c4d83d5-56f7-4eed-b897-6368f9067750}",
										"_name": "Pokaż symbol gdy ryzyka istnieją",
										"_idname": "DISPLAY_SYMBOLS_IF_RISKS_OR_CONTROLS",
										"_type": "BOOL",
										"_controltype": "BOOL"
									}
								],
								"_name": "Wizualizacja atrybutów"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{639f3143-2078-4696-b698-963efac30971}",
				"_name": "Zadanie B",
				"_class": "Zadanie",
				"_idclass": "C_TASK",
				"_repoid": "{04accc91-0388-4fbf-bef5-a250d9bbb023}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zadanie E"
										},
										"_uuid": "{bd5a34a0-6a9f-4e8b-9961-9175e0e6d680}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "07",
										"_uuid": "{21c6687e-01ee-4e68-bc16-b908d5e2c4e0}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"_name": "Diagram odpowiedzialności (RACI)"
						},
						{
							"_name": "Zasoby"
						},
						{
							"group": {
								"_name": "Wejście/Wyjście (w trakcie wykonywania)"
							},
							"_name": "Wejście/Wyjście"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Ciągłe usprawnianie"
						},
						{
							"_name": "Zarządzanie ryzykiem"
						},
						{
							"_name": "Kontrole"
						},
						{
							"_name": "Systemy/Produkty"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Powiązane obiekty"
								},
								{
									"_name": "Definicja obiektu"
								},
								{
									"_name": "Szczegóły importu"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"_name": "Pętla - standardowa"
						},
						{
							"group": {
								"_name": "Wywoływanie zdarzeń"
							},
							"_name": "Pętla - wiele instancji"
						},
						{
							"group": [
								{
									"_name": "Szczegóły implementacji"
								},
								{
									"_name": "Usługa (szczegóły)"
								},
								{
									"_name": "Wysłanie/Odbiór (szczegóły)"
								},
								{
									"_name": "Użytkownika (szczegóły)"
								},
								{
									"_name": "Skrypt (szczegóły)"
								},
								{
									"_name": "Wywołanie aktywności (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"_name": "Dane symulacyjne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{6e1ca5d9-2b44-435e-90cc-75a69c799157}",
									"_value": "v0",
									"_name": "w środku",
									"_constraint": "v0@v1",
									"_nameconstraint": "w środku@poniżej"
								},
								"_id": "{565e3a93-18d2-4518-bbb3-2dcd433e85d0}",
								"_name": "Pokaż nazwę",
								"_idname": "REPRESENTATION_NAME_TASK",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"group": {
								"attribute": [
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{4caf788d-c93c-48b8-a8b8-c7510cbe7718}"
										},
										"_id": "{7abda636-e481-4d1f-8a8f-debaa7ed8394}",
										"_name": "Pokaż role odpowiedzialne za wykonanie",
										"_idname": "DISPLAY_RESPONSIBLE_FOR_EXECUTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{d38bab97-c4e8-4b1b-9047-3bf98b0f8067}"
										},
										"_id": "{22c826d9-af67-49a0-8cc5-c536b4e9caa9}",
										"_name": "Pokaż symbol (\"i\"), jeżeli istnieje opis",
										"_idname": "DISPLAY_SYMBOL_IF_DESCRIPTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{15c9ff35-f72f-4e70-a766-b7597d2fa279}"
										},
										"_id": "{0c4d83d5-56f7-4eed-b897-6368f9067750}",
										"_name": "Pokaż symbol gdy ryzyka istnieją",
										"_idname": "DISPLAY_SYMBOLS_IF_RISKS_OR_CONTROLS",
										"_type": "BOOL",
										"_controltype": "BOOL"
									}
								],
								"_name": "Wizualizacja atrybutów"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{a3d8639e-9dde-472b-b920-b316ed5586fc}",
				"_name": "Zadanie E",
				"_class": "Zadanie",
				"_idclass": "C_TASK",
				"_repoid": "{04accc91-0388-4fbf-bef5-a250d9bbb023}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zadanie D"
										},
										"_uuid": "{93012ac7-363f-4351-b07b-a044f193cb08}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "06",
										"_uuid": "{065c2bf6-c336-4e2f-aaeb-85f69fa32f60}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"_name": "Diagram odpowiedzialności (RACI)"
						},
						{
							"_name": "Zasoby"
						},
						{
							"group": {
								"_name": "Wejście/Wyjście (w trakcie wykonywania)"
							},
							"_name": "Wejście/Wyjście"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Ciągłe usprawnianie"
						},
						{
							"_name": "Zarządzanie ryzykiem"
						},
						{
							"_name": "Kontrole"
						},
						{
							"_name": "Systemy/Produkty"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Powiązane obiekty"
								},
								{
									"_name": "Definicja obiektu"
								},
								{
									"_name": "Szczegóły importu"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"_name": "Pętla - standardowa"
						},
						{
							"group": {
								"_name": "Wywoływanie zdarzeń"
							},
							"_name": "Pętla - wiele instancji"
						},
						{
							"group": [
								{
									"_name": "Szczegóły implementacji"
								},
								{
									"_name": "Usługa (szczegóły)"
								},
								{
									"_name": "Wysłanie/Odbiór (szczegóły)"
								},
								{
									"_name": "Użytkownika (szczegóły)"
								},
								{
									"_name": "Skrypt (szczegóły)"
								},
								{
									"_name": "Wywołanie aktywności (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"_name": "Dane symulacyjne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{6e1ca5d9-2b44-435e-90cc-75a69c799157}",
									"_value": "v0",
									"_name": "w środku",
									"_constraint": "v0@v1",
									"_nameconstraint": "w środku@poniżej"
								},
								"_id": "{565e3a93-18d2-4518-bbb3-2dcd433e85d0}",
								"_name": "Pokaż nazwę",
								"_idname": "REPRESENTATION_NAME_TASK",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"group": {
								"attribute": [
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{4caf788d-c93c-48b8-a8b8-c7510cbe7718}"
										},
										"_id": "{7abda636-e481-4d1f-8a8f-debaa7ed8394}",
										"_name": "Pokaż role odpowiedzialne za wykonanie",
										"_idname": "DISPLAY_RESPONSIBLE_FOR_EXECUTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{d38bab97-c4e8-4b1b-9047-3bf98b0f8067}"
										},
										"_id": "{22c826d9-af67-49a0-8cc5-c536b4e9caa9}",
										"_name": "Pokaż symbol (\"i\"), jeżeli istnieje opis",
										"_idname": "DISPLAY_SYMBOL_IF_DESCRIPTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{15c9ff35-f72f-4e70-a766-b7597d2fa279}"
										},
										"_id": "{0c4d83d5-56f7-4eed-b897-6368f9067750}",
										"_name": "Pokaż symbol gdy ryzyka istnieją",
										"_idname": "DISPLAY_SYMBOLS_IF_RISKS_OR_CONTROLS",
										"_type": "BOOL",
										"_controltype": "BOOL"
									}
								],
								"_name": "Wizualizacja atrybutów"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{ab21002b-a6f9-4a81-b923-a1173cc2245d}",
				"_name": "Zadanie D",
				"_class": "Zadanie",
				"_idclass": "C_TASK",
				"_repoid": "{04accc91-0388-4fbf-bef5-a250d9bbb023}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zadanie A"
										},
										"_uuid": "{85da728d-71be-4615-912c-931c5e2b0dc5}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "02",
										"_uuid": "{4e8dac06-8b39-4542-b2b2-338c3c39137b}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{bcfca9dc-c54f-4b73-9280-38b8ed8a7b10}",
											"_type": "LONGSTRING"
										},
										"value": "Opis A",
										"richvalue": {
											"x": "Opis A"
										},
										"_uuid": "{d4a0bf24-18e5-4b01-a02d-272e22359404}"
									},
									"_id": "{2fcdce9b-4a60-4011-a3fb-b435ba8d2a6a}",
									"_name": "Opis",
									"_idname": "A_DESCRIPTION_CS",
									"_type": "LONGSTRING",
									"_controltype": "STRING_MULTILINE"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"relation": [
								{
									"link": {
										"endpoint": {
											"_type": "repoinst",
											"_role": "TO",
											"_objectid": "{67f0fd50-90eb-4f5d-92e1-577fffeda9e9}",
											"_name": "Rola A",
											"_class": "Rola"
										},
										"_id": "{52886154-e073-44d0-90bf-4272798894ee}"
									},
									"_class": "Odpowiedzialny za wykonanie",
									"_idclass": "RESPONSIBLE_FOR_EXECUTION",
									"_label": "Odpowiedzialny za wykonanie",
									"_controltype": "RELATIONS"
								},
								{
									"link": {
										"endpoint": {
											"_type": "repoinst",
											"_role": "TO",
											"_objectid": "{f97cab4c-95a7-4b65-9dd3-7322898bae86}",
											"_name": "Rola C",
											"_class": "Rola"
										},
										"_id": "{802e24c6-a17a-49ab-9ba6-05061dae0930}"
									},
									"_class": "Informowany",
									"_idclass": "TO_INFORM",
									"_label": "Informowany",
									"_controltype": "RELATIONS"
								}
							],
							"_name": "Diagram odpowiedzialności (RACI)"
						},
						{
							"relation": [
								{
									"link": {
										"endpoint": {
											"_type": "repoinst",
											"_role": "TO",
											"_objectid": "{d353e2c5-09ac-4ee9-b59e-510a30216f50}",
											"_name": "Lista kontrolna ABC ",
											"_class": "Dokument"
										},
										"_id": "{b61eab65-9aab-4f45-bea4-23e2630bf177}"
									},
									"_class": "Powiązane dokumenty",
									"_idclass": "RC_REFERENCED_DOCUMENTS_M",
									"_label": "Powiązane dokumenty",
									"_controltype": "RELATIONS"
								},
								{
									"link": {
										"endpoint": {
											"_type": "repoinst",
											"_role": "TO",
											"_objectid": "{0b60b209-391a-4e7d-8f92-373ebc8c5b9a}",
											"_name": "Aplikacja A",
											"_class": "Aplikacja"
										},
										"_id": "{1b0ead27-0fbe-42b0-bcf8-e0c7e1bebd32}"
									},
									"_class": "Wsparcie IT",
									"_idclass": "RC_REFERENCED_IT_SYSTEM_ELEMENTS_CS",
									"_label": "Wsparcie IT",
									"_controltype": "RELATIONS"
								}
							],
							"_name": "Zasoby"
						},
						{
							"relation": [
								{
									"link": {
										"endpoint": {
											"_type": "repoinst",
											"_role": "TO",
											"_objectid": "{c0e6cee6-3ad7-4c84-9fe9-f7573fbeafe0}",
											"_name": "Dokument A ",
											"_class": "Dokument"
										},
										"_id": "{f83141a5-e993-4dac-9057-6685fdb2a134}"
									},
									"_class": "Dane wejściowe (opcjonalne)",
									"_idclass": "RC_REFERENCED_INPUT_DATA_OPTIONAL_M",
									"_label": "Dane wejściowe (opcjonalne)",
									"_controltype": "RELATIONS"
								},
								{
									"link": {
										"endpoint": {
											"_type": "repoinst",
											"_role": "TO",
											"_objectid": "{ce0dd4e5-8607-44bb-a391-5966b8a088f1}",
											"_name": "Dokument B ",
											"_class": "Dokument"
										},
										"_id": "{e738c9bc-abf2-4c65-b893-6dd16f7b8670}"
									},
									"_class": "Dane wyjściowe (opcjonalne)",
									"_idclass": "RC_REFERENCED_OUTPUT_DATA_OPTIONAL_M",
									"_label": "Dane wyjściowe (opcjonalne)",
									"_controltype": "RELATIONS"
								}
							],
							"group": {
								"_name": "Wejście/Wyjście (w trakcie wykonywania)"
							},
							"_name": "Wejście/Wyjście"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Ciągłe usprawnianie"
						},
						{
							"_name": "Zarządzanie ryzykiem"
						},
						{
							"_name": "Kontrole"
						},
						{
							"_name": "Systemy/Produkty"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Powiązane obiekty"
								},
								{
									"_name": "Definicja obiektu"
								},
								{
									"_name": "Szczegóły importu"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"_name": "Pętla - standardowa"
						},
						{
							"group": {
								"_name": "Wywoływanie zdarzeń"
							},
							"_name": "Pętla - wiele instancji"
						},
						{
							"group": [
								{
									"_name": "Szczegóły implementacji"
								},
								{
									"_name": "Usługa (szczegóły)"
								},
								{
									"_name": "Wysłanie/Odbiór (szczegóły)"
								},
								{
									"_name": "Użytkownika (szczegóły)"
								},
								{
									"_name": "Skrypt (szczegóły)"
								},
								{
									"_name": "Wywołanie aktywności (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"_name": "Dane symulacyjne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{6e1ca5d9-2b44-435e-90cc-75a69c799157}",
									"_value": "v0",
									"_name": "w środku",
									"_constraint": "v0@v1",
									"_nameconstraint": "w środku@poniżej"
								},
								"_id": "{565e3a93-18d2-4518-bbb3-2dcd433e85d0}",
								"_name": "Pokaż nazwę",
								"_idname": "REPRESENTATION_NAME_TASK",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"group": {
								"attribute": [
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{4caf788d-c93c-48b8-a8b8-c7510cbe7718}"
										},
										"_id": "{7abda636-e481-4d1f-8a8f-debaa7ed8394}",
										"_name": "Pokaż role odpowiedzialne za wykonanie",
										"_idname": "DISPLAY_RESPONSIBLE_FOR_EXECUTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{d38bab97-c4e8-4b1b-9047-3bf98b0f8067}"
										},
										"_id": "{22c826d9-af67-49a0-8cc5-c536b4e9caa9}",
										"_name": "Pokaż symbol (\"i\"), jeżeli istnieje opis",
										"_idname": "DISPLAY_SYMBOL_IF_DESCRIPTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{15c9ff35-f72f-4e70-a766-b7597d2fa279}"
										},
										"_id": "{0c4d83d5-56f7-4eed-b897-6368f9067750}",
										"_name": "Pokaż symbol gdy ryzyka istnieją",
										"_idname": "DISPLAY_SYMBOLS_IF_RISKS_OR_CONTROLS",
										"_type": "BOOL",
										"_controltype": "BOOL"
									}
								],
								"_name": "Wizualizacja atrybutów"
							},
							"_name": "Postać"
						}
					]
				},
				"_id": "{bb656ae6-cb46-4fcb-a470-659dfed3f572}",
				"_name": "Zadanie A",
				"_class": "Zadanie",
				"_idclass": "C_TASK",
				"_repoid": "{04accc91-0388-4fbf-bef5-a250d9bbb023}"
			},
			{
				"notebook": {
					"chapter": [
						{
							"attribute": [
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{f999485c-3399-4c55-8825-317ddec741f6}",
											"_type": "ADOSTRING"
										},
										"value": {
											"p": "Zadanie C"
										},
										"_uuid": "{bee7d034-335f-4d4b-9da5-db29f1da439b}"
									},
									"_id": "{16656faa-147a-4a46-917b-108016d56c0d}",
									"_name": "Nazwa",
									"_idname": "NAME",
									"_type": "ADOSTRING",
									"_controltype": "NAME"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{22bced4d-ef40-40d6-adff-f4ca8b27b99b}",
											"_type": "SHORTSTRING"
										},
										"value": "04",
										"_uuid": "{6cec2a3c-3cc0-41cb-b6ff-3bd50e045d55}"
									},
									"_id": "{873bea07-6e9a-4c30-b9e3-64212ca8aad7}",
									"_name": "Kolejność",
									"_idname": "A_ORDER",
									"_type": "SHORTSTRING",
									"_controltype": "STRING"
								},
								{
									"attrval": {
										"attrvaltype": {
											"_uuid": "{bcfca9dc-c54f-4b73-9280-38b8ed8a7b10}",
											"_type": "LONGSTRING"
										},
										"value": "Opis C",
										"richvalue": {
											"x": "Opis C"
										},
										"_uuid": "{e1893187-83c0-49bb-abf8-338e3104db39}"
									},
									"_id": "{2fcdce9b-4a60-4011-a3fb-b435ba8d2a6a}",
									"_name": "Opis",
									"_idname": "A_DESCRIPTION_CS",
									"_type": "LONGSTRING",
									"_controltype": "STRING_MULTILINE"
								}
							],
							"_name": "Informacje ogólne"
						},
						{
							"relation": {
								"link": {
									"endpoint": {
										"_type": "repoinst",
										"_role": "TO",
										"_objectid": "{f97cab4c-95a7-4b65-9dd3-7322898bae86}",
										"_name": "Rola C",
										"_class": "Rola"
									},
									"_id": "{8cc0e87a-c710-481b-b861-98bfc1895a6e}"
								},
								"_class": "Odpowiedzialny za wykonanie",
								"_idclass": "RESPONSIBLE_FOR_EXECUTION",
								"_label": "Odpowiedzialny za wykonanie",
								"_controltype": "RELATIONS"
							},
							"_name": "Diagram odpowiedzialności (RACI)"
						},
						{
							"_name": "Zasoby"
						},
						{
							"group": {
								"_name": "Wejście/Wyjście (w trakcie wykonywania)"
							},
							"_name": "Wejście/Wyjście"
						},
						{
							"group": {
								"_name": "Wymagana akcja"
							},
							"_name": "Ciągłe usprawnianie"
						},
						{
							"_name": "Zarządzanie ryzykiem"
						},
						{
							"_name": "Kontrole"
						},
						{
							"_name": "Systemy/Produkty"
						},
						{
							"group": [
								{
									"_name": "Zastosowanie obiektu"
								},
								{
									"_name": "Powiązane obiekty"
								},
								{
									"_name": "Definicja obiektu"
								},
								{
									"_name": "Szczegóły importu"
								}
							],
							"_name": "Właściwości obiektu"
						},
						{
							"_name": "Pętla - standardowa"
						},
						{
							"group": {
								"_name": "Wywoływanie zdarzeń"
							},
							"_name": "Pętla - wiele instancji"
						},
						{
							"group": [
								{
									"_name": "Szczegóły implementacji"
								},
								{
									"_name": "Usługa (szczegóły)"
								},
								{
									"_name": "Wysłanie/Odbiór (szczegóły)"
								},
								{
									"_name": "Użytkownika (szczegóły)"
								},
								{
									"_name": "Skrypt (szczegóły)"
								},
								{
									"_name": "Wywołanie aktywności (szczegóły)"
								}
							],
							"_name": "Szczegóły"
						},
						{
							"_name": "Zmiany"
						},
						{
							"_name": "Załączniki"
						},
						{
							"_name": "Dane symulacyjne"
						},
						{
							"attribute": {
								"attrval": {
									"attrvaltype": {
										"_uuid": "{6c286bc9-9ea0-4b0b-8c4e-9b1837a4ec7e}",
										"_type": "ENUM"
									},
									"_uuid": "{6e1ca5d9-2b44-435e-90cc-75a69c799157}",
									"_value": "v0",
									"_name": "w środku",
									"_constraint": "v0@v1",
									"_nameconstraint": "w środku@poniżej"
								},
								"_id": "{565e3a93-18d2-4518-bbb3-2dcd433e85d0}",
								"_name": "Pokaż nazwę",
								"_idname": "REPRESENTATION_NAME_TASK",
								"_type": "ENUM",
								"_controltype": "ENUM"
							},
							"group": {
								"attribute": [
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{4caf788d-c93c-48b8-a8b8-c7510cbe7718}"
										},
										"_id": "{7abda636-e481-4d1f-8a8f-debaa7ed8394}",
										"_name": "Pokaż role odpowiedzialne za wykonanie",
										"_idname": "DISPLAY_RESPONSIBLE_FOR_EXECUTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{d38bab97-c4e8-4b1b-9047-3bf98b0f8067}"
										},
										"_id": "{22c826d9-af67-49a0-8cc5-c536b4e9caa9}",
										"_name": "Pokaż symbol (\"i\"), jeżeli istnieje opis",
										"_idname": "DISPLAY_SYMBOL_IF_DESCRIPTION",
										"_type": "BOOL",
										"_controltype": "BOOL"
									},
									{
										"attrval": {
											"attrvaltype": {
												"_uuid": "{fece26d2-b1ad-441e-81ef-e237b3204b8f}",
												"_type": "BOOL"
											},
											"value": "1",
											"_uuid": "{15c9ff35-f72f-4e70-a766-b7597d2fa279}"
										},
										"_id": "{0c4d83d5-56f7-4eed-b897-6368f9067750}",
										"_name": "Pokaż symbol gdy ryzyka istnieją",
										"_idname": "DISPLAY_SYMBOLS_IF_RISKS_OR_CONTROLS",
										"_type": "BOOL",
										"_controltype": "BOOL"
									}
								],
								"_name": "Wizualizacja atrybutów"
							},
							"_name": "Postać"
						}
					],
					"__text": "`"
				},
				"_id": "{f29109e6-f09f-43d9-a2c1-3dd3dcedc03e}",
				"_name": "Zadanie C",
				"_class": "Zadanie",
				"_idclass": "C_TASK",
				"_repoid": "{04accc91-0388-4fbf-bef5-a250d9bbb023}"
			}
		],
		"image": {
			"_scale": "-1.00",
			"_width": "18465",
			"_height": "5685",
			"_href": "model1132d8c5-e71f-45ec-ab6a-3162ca99bc81(-1.00)0.svg"
		},
		"_id": "{bd41281b-a0dd-4743-87aa-00a006d081c2}",
		"_name": "Przykładowy diagram IC (wersjonowanie)",
		"_class": "Diagram procesów biznesowych (BPMN 2.0)",
		"_idclass": "MT_BUSINESS_PROCESS_DIAGRAM_BPMN_20",
		"_variantid": "{1132d8c5-e71f-45ec-ab6a-3162ca99bc81}",
		"_lang": "pl",
		"base64image":"data:image/png;base64,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jhwZFRWV4jl53b59e48ePerVq6fxVatW7du376lTpxxj9u7d27lz58GDB+vPR48eVQW6Bu/evVs/6tN8Z8+0c8eu9PCePXtWTFOlShU/P7+wsDDHmHXr1umx48ePT0lb7FaH0C+i8fXr19eyz8nJv5uEc8qUKRo8Z84cx07i4+N1A0utWrXMsVq1arV582ZeOQAAAAAAwFeRvAIA4MXixYv//Oc/69/KgQMHum49dOhQmzZtlCZUr15docnx48dTPCevJ0+eHDRoUPPmzU0A0bFjR9c0JCYmxgQlymi0uGSXLl00cvLkySm/hSCezJgxw7Er7UFZjxIcczjtav/+/Y4xwcHBeqw2XdLygWk5UePGjc2vo8fqCdgHKy7RYE2t6noeVq5cqYVyzYGUNM2ePdsRysA7ktc8SS/0//u//9O73XVxU72jatSo4Zis/IknntAlwzV5VeFshQoV/vSnPznG/+1vf3PseerUqWax5RMnTtx9991m2Lx587Tpo48+8jJPuv36pefWrl27m266yTHmzjvvVERqP1zXrl3V//rrr+twL730kmP8t99+ay4cRtGiRdVZqlQp+x62bdv25JNPuj6fH3/88eJFVhEDAAAAAAA+iOQVAABPgoKC7r33Xv1D+fPPP7tuVSWYkhF7mqC62FmzZrlNXocMGWICF4dKlSrZQ8qIiAjTv2fPnv/85z/mzyow06YOHTp4CVaKFy9uP5wK0v75z386xmgxStXI2Q+nZWvNpnPnzv3000+O8U899VRoaKg1uFevXib0sR8oPDz8008/dX0+L7/8smrqeAllEMlr3hMYGKioUq/1cuXKuW5VWap5J9SsWTMgICAkJGT58uWffPKJ8su//vWvjuTVlNUrxNUfFFUePnxYiyp/9tlnpnPDhg3WSFWUqvMf//iHktoHHnhAO69Tp46pedWmwS7Mm1OZrqZTt3aimz7MnvXYgwcP6k2+du3aEiVKmM5ly5ZZI1V3q85//etf77777muvvab7O3TJWLFixZdffml+OyXB1mDX5PXYsWN33HGHOgsVKrRgwQJdZfbt26d7N26//XZ16oi8igAAAAAAgO8heQUAwC3NkfnGG2+Y3MFR/ZmSNuunqRn74IMP1qxZo3k6t27dqiDjtttu04KPjuRVmYspjVORmMpDFawoTNGsoqZTJafWyNjYWJNoNGjQ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6j4IvW0aN258991335Dm4Ycfbt26tV7uru/zHj16FChQ4Ibf3H777d99951mD3fdeWhoaL169e677z4z8pZbbtHn702bNlkDjh49eqdXSm2twUoo/fz8Xn75ZS0NbXaoPyheXbNmjTVGO9ejzIC//OUvZidz5sxJ8Zy8jhkzRmM0O/GUKVOszuDg4IoVK957773Wr/nII480adIkKuryveqHDh0yO9evYN9bXFzcAw88oH6dWMeBhg4dqn49f/tRKlSocM8991hHKVSokGLjLIyo8+rLsuO4DK7z+kqt/hcuXr4WK2B4sVofT1/DjVm23e2xTPLqpf3Syc8qa5B8RWuo0yo/tWjKTfUrXk3xsM6r5XR4pAIJNRPr7joS5P1slGjZX8MmLtlg9egrwvv+V01T0jm+1wOQd+07EZrxdV57TFttdWraYU/XPRWEeTqcSV49tQeKVrcvRD124VozsbBjJ4dPhqhfV0Xzo+s6r1f+cUxIPHb6rK57upRpjNbJ9nIqTPKq+TkLlmximla51mztyoNVZdt5zOz4P/y/koDP0ATCGVzn9any3YPDLn8O161sHzcb5unS1268+wUmTPLqpWmCdPsnK7Pwqm6Acz7n/v7qbzv88oKsruu8WlRHaz7ymfv27HOfOGhFCQ3QnnlJAPANJK8AAHjiKXlt2LChsoOHHnpox44dVueuXbu++eYbLdpqsoM//elP//rXv/r16+eID2rVqqXH6r/6swKUp556yoy/+eabv/jiC0d+IYo2+/btqxmJzbA///nPr732mr+/m/8fmTdv3jvvvHPjjTeakfrD22+/PXv2bGtA06ZNdehixYo5HjhixAiTnmgPjk3ag/r79Olj9cydO/c///mPdZS///3vOj/btm3j1ZIZJK85rGjRono1K//79NNP9dZVLqgs87HHHjOv8g8++MCefSqd1Vb1/9///d9bb71VtmxZvfO1VrNJVU3AaTlx4sTjjz+uTTfddJPen3pH3XbbbSYQnTr18vJLnpJXjTGXEvvRf/rpJ/OsNFP5119/XbhwYZNZ6tIwc+ZM62Kk8NLkx4qEC6VZtiz1+xe3yevChQvNsXr16mV1BgQE3H///WYPn3/+efny5T/88EOT5j777LOKS80whdPqGTlypH2Hy5cvtwJpxxXQXFUrV65sfty5c6fJpHXqdJQyZcq88cYb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"	}
}