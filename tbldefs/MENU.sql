CREATE TABLE [MENU] (
  [Men_Page] LONG,
  [Men_Numero] SHORT,
  [Men_Libelle] VARCHAR (255),
  [Men_Command] SHORT,
  [Men_Argument] VARCHAR (50),
  [Men_Argument1] VARCHAR (50),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Men_Page], [Men_Numero])
)
