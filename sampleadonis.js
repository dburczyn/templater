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
		"_lang": "pl"
	}
}