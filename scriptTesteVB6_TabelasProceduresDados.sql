USE [TESTE_VB6]
GO
/****** Object:  Table [dbo].[Cardapio]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Cardapio](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[IdCozinha] [int] NOT NULL,
	[TempoPreparo] [int] NOT NULL,
	[Preco] [decimal](9, 2) NOT NULL,
	[Descricao] [nchar](100) NULL,
 CONSTRAINT [PK_Cardapio] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Comandas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Comandas](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[IdTipoComanda] [int] NOT NULL,
	[NumeroMesa] [int] NOT NULL,
	[QuantidadePessoa] [int] NOT NULL,
	[DataConfirmacaoPreparo] [datetime] NOT NULL,
	[DataPrevistaPreparo] [datetime] NULL,
	[DataFinalizacaoPreparo] [datetime] NULL,
	[StatusComanda] [int] NOT NULL,
 CONSTRAINT [PK_Comandas] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ComandasItem]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ComandasItem](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[IdComanda] [int] NOT NULL,
	[IdCardapio] [int] NOT NULL,
	[Quantidade] [int] NOT NULL,
	[Preco] [decimal](9, 2) NOT NULL,
	[TotalPreco] [decimal](9, 2) NOT NULL,
	[DataConfirmacaoPreparo] [datetime] NOT NULL,
	[DataPrevistaPreparo] [datetime] NULL,
	[DataFinalizacaoPreparo] [datetime] NULL,
	[StatusItem] [int] NOT NULL,
 CONSTRAINT [PK_ComandasItem] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Cozinhas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Cozinhas](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Descricao] [nvarchar](100) NOT NULL,
	[Capacidade] [int] NOT NULL,
 CONSTRAINT [PK_Cozinhas] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TiposComandas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TiposComandas](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Descricao] [nvarchar](100) NOT NULL,
 CONSTRAINT [PK_TiposComandas] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[Cardapio] ON 

INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (1, 8, 3, CAST(22.50 AS Decimal(9, 2)), N'SALADA CAESAR                                                                                       ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (2, 8, 22, CAST(80.00 AS Decimal(9, 2)), N'FILET MIGNON À PARMEGIANA                                                                           ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (3, 8, 25, CAST(44.80 AS Decimal(9, 2)), N'ANCHO GRELHADO                                                                                      ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (4, 8, 30, CAST(40.00 AS Decimal(9, 2)), N'NHOQUE DE BATATA                                                                                    ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (5, 13, 3, CAST(12.00 AS Decimal(9, 2)), N'PUDIM DE LEITE                                                                                      ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (6, 13, 3, CAST(18.00 AS Decimal(9, 2)), N'TORTA BANOFFEE (FATIA)                                                                              ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (7, 9, 5, CAST(8.00 AS Decimal(9, 2)), N'SUCO DE LARANJA                                                                                     ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (8, 9, 5, CAST(8.00 AS Decimal(9, 2)), N'SUCO DE LIMAO                                                                                       ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (9, 9, 5, CAST(16.00 AS Decimal(9, 2)), N'SUCO HEART DETOX 350 ML                                                                             ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (10, 9, 5, CAST(122.50 AS Decimal(9, 2)), N'VINHO LEYDA PINOT NOIR 2018 750ML                                                                   ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (11, 15, 7, CAST(60.00 AS Decimal(9, 2)), N'PIZZA CALABRESA                                                                                     ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (12, 15, 7, CAST(65.00 AS Decimal(9, 2)), N'PIZZA ALLICI                                                                                        ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (13, 15, 7, CAST(65.00 AS Decimal(9, 2)), N'PIZZA PEPPERONE                                                                                     ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (14, 15, 18, CAST(70.00 AS Decimal(9, 2)), N'PIZZA PIZZAIOLA                                                                                     ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (15, 15, 18, CAST(70.00 AS Decimal(9, 2)), N'PIZZA CALZONE DE COGUMELOS COM MUÇARELA                                                             ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (16, 10, 3, CAST(18.90 AS Decimal(9, 2)), N'CROCANTE DE SALMAO 2 UNIDADES                                                                       ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (17, 10, 3, CAST(22.00 AS Decimal(9, 2)), N'CEVICHE                                                                                             ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (18, 10, 3, CAST(20.00 AS Decimal(9, 2)), N'SUNOMONO                                                                                            ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (19, 10, 3, CAST(40.00 AS Decimal(9, 2)), N'TAKO PESTO 7 UNIDADES                                                                               ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (20, 10, 3, CAST(40.00 AS Decimal(9, 2)), N'TEKAMAKI 8 UNIDADES                                                                                 ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (21, 12, 12, CAST(33.00 AS Decimal(9, 2)), N'EDAMAME                                                                                             ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (22, 12, 12, CAST(44.00 AS Decimal(9, 2)), N'ARROZ CRISPY                                                                                        ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (23, 12, 12, CAST(33.00 AS Decimal(9, 2)), N'HARUMAKI DE LEGUMES 2 UNIDADES                                                                      ')
INSERT [dbo].[Cardapio] ([Id], [IdCozinha], [TempoPreparo], [Preco], [Descricao]) VALUES (24, 12, 33, CAST(65.00 AS Decimal(9, 2)), N'YAKISSOBA DE CARNE                                                                                  ')
SET IDENTITY_INSERT [dbo].[Cardapio] OFF
GO
SET IDENTITY_INSERT [dbo].[Comandas] ON 

INSERT [dbo].[Comandas] ([Id], [IdTipoComanda], [NumeroMesa], [QuantidadePessoa], [DataConfirmacaoPreparo], [DataPrevistaPreparo], [DataFinalizacaoPreparo], [StatusComanda]) VALUES (1, 1, 12, 2, CAST(N'2023-10-17T16:01:57.270' AS DateTime), CAST(N'2023-10-17T16:28:35.000' AS DateTime), CAST(N'2023-10-17T16:09:24.373' AS DateTime), 2)
SET IDENTITY_INSERT [dbo].[Comandas] OFF
GO
SET IDENTITY_INSERT [dbo].[ComandasItem] ON 

INSERT [dbo].[ComandasItem] ([Id], [IdComanda], [IdCardapio], [Quantidade], [Preco], [TotalPreco], [DataConfirmacaoPreparo], [DataPrevistaPreparo], [DataFinalizacaoPreparo], [StatusItem]) VALUES (1, 1, 2, 1, CAST(80.00 AS Decimal(9, 2)), CAST(80.00 AS Decimal(9, 2)), CAST(N'2023-10-17T16:03:35.000' AS DateTime), CAST(N'2023-10-17T16:25:35.000' AS DateTime), CAST(N'2023-10-17T16:07:50.337' AS DateTime), 4)
INSERT [dbo].[ComandasItem] ([Id], [IdComanda], [IdCardapio], [Quantidade], [Preco], [TotalPreco], [DataConfirmacaoPreparo], [DataPrevistaPreparo], [DataFinalizacaoPreparo], [StatusItem]) VALUES (2, 1, 1, 2, CAST(22.50 AS Decimal(9, 2)), CAST(45.00 AS Decimal(9, 2)), CAST(N'2023-10-17T16:03:35.000' AS DateTime), CAST(N'2023-10-17T16:06:35.000' AS DateTime), CAST(N'2023-10-17T16:07:10.397' AS DateTime), 4)
INSERT [dbo].[ComandasItem] ([Id], [IdComanda], [IdCardapio], [Quantidade], [Preco], [TotalPreco], [DataConfirmacaoPreparo], [DataPrevistaPreparo], [DataFinalizacaoPreparo], [StatusItem]) VALUES (3, 1, 7, 2, CAST(8.00 AS Decimal(9, 2)), CAST(16.00 AS Decimal(9, 2)), CAST(N'2023-10-17T16:03:35.000' AS DateTime), CAST(N'2023-10-17T16:08:35.000' AS DateTime), NULL, 5)
INSERT [dbo].[ComandasItem] ([Id], [IdComanda], [IdCardapio], [Quantidade], [Preco], [TotalPreco], [DataConfirmacaoPreparo], [DataPrevistaPreparo], [DataFinalizacaoPreparo], [StatusItem]) VALUES (4, 1, 3, 1, CAST(44.80 AS Decimal(9, 2)), CAST(44.80 AS Decimal(9, 2)), CAST(N'2023-10-17T16:03:35.000' AS DateTime), CAST(N'2023-10-17T16:28:35.000' AS DateTime), CAST(N'2023-10-17T16:07:59.487' AS DateTime), 4)
SET IDENTITY_INSERT [dbo].[ComandasItem] OFF
GO
SET IDENTITY_INSERT [dbo].[Cozinhas] ON 

INSERT [dbo].[Cozinhas] ([Id], [Descricao], [Capacidade]) VALUES (8, N'COZINHA PRINCIPAL ', 2)
INSERT [dbo].[Cozinhas] ([Id], [Descricao], [Capacidade]) VALUES (9, N'BAR', 1)
INSERT [dbo].[Cozinhas] ([Id], [Descricao], [Capacidade]) VALUES (10, N'SUSHI FRIO', 2)
INSERT [dbo].[Cozinhas] ([Id], [Descricao], [Capacidade]) VALUES (12, N'SUSHI QUENTE', 2)
INSERT [dbo].[Cozinhas] ([Id], [Descricao], [Capacidade]) VALUES (13, N'SOBREMESAS', 1)
INSERT [dbo].[Cozinhas] ([Id], [Descricao], [Capacidade]) VALUES (15, N'PIZZARIA', 2)
SET IDENTITY_INSERT [dbo].[Cozinhas] OFF
GO
SET IDENTITY_INSERT [dbo].[TiposComandas] ON 

INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (1, N'Comanda 01')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (2, N'Comanda 02')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (3, N'Comanda 03')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (4, N'Comanda 04')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (5, N'Comanda 05')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (10, N'Comanda 10')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (11, N'Comanda 11')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (12, N'Comanda 12')
INSERT [dbo].[TiposComandas] ([Id], [Descricao]) VALUES (13, N'Comanda 13')
SET IDENTITY_INSERT [dbo].[TiposComandas] OFF
GO
ALTER TABLE [dbo].[Cardapio]  WITH CHECK ADD  CONSTRAINT [FK_Cardapio_Cozinhas_IdCozinha] FOREIGN KEY([IdCozinha])
REFERENCES [dbo].[Cozinhas] ([Id])
GO
ALTER TABLE [dbo].[Cardapio] CHECK CONSTRAINT [FK_Cardapio_Cozinhas_IdCozinha]
GO
ALTER TABLE [dbo].[Comandas]  WITH CHECK ADD  CONSTRAINT [FK_Comandas_TiposComandas_IdTipoComanda] FOREIGN KEY([IdTipoComanda])
REFERENCES [dbo].[TiposComandas] ([Id])
GO
ALTER TABLE [dbo].[Comandas] CHECK CONSTRAINT [FK_Comandas_TiposComandas_IdTipoComanda]
GO
ALTER TABLE [dbo].[ComandasItem]  WITH CHECK ADD  CONSTRAINT [FK_ComandasItem_Cardapio_IdCardapio] FOREIGN KEY([IdCardapio])
REFERENCES [dbo].[Cardapio] ([Id])
GO
ALTER TABLE [dbo].[ComandasItem] CHECK CONSTRAINT [FK_ComandasItem_Cardapio_IdCardapio]
GO
ALTER TABLE [dbo].[ComandasItem]  WITH CHECK ADD  CONSTRAINT [FK_ComandasItem_Comandas_IdComanda] FOREIGN KEY([IdComanda])
REFERENCES [dbo].[Comandas] ([Id])
GO
ALTER TABLE [dbo].[ComandasItem] CHECK CONSTRAINT [FK_ComandasItem_Comandas_IdComanda]
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_D_Cardapio]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




CREATE PROCEDURE [dbo].[SP_Teste_D_Cardapio](@Id INT)
AS
BEGIN
    
    DELETE FROM Cardapio 
    WHERE Id = @Id;
         
END 
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_D_ComandaItens]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[SP_Teste_D_ComandaItens](@Id INT)
AS
BEGIN
    
    DELETE FROM ComandasItem
    WHERE Id = @Id;
         
END 
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_D_Cozinhas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[SP_Teste_D_Cozinhas](@Id INT)
AS
BEGIN
    
    DELETE FROM Cozinhas 
    WHERE Id = @Id;
         
END 
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_D_TiposComandas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




CREATE PROCEDURE [dbo].[SP_Teste_D_TiposComandas](@Id INT)
AS
BEGIN
    
    DELETE FROM TiposComandas 
    WHERE Id = @Id;
         
END 
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_I_Cardapio]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






CREATE PROCEDURE [dbo].[SP_Teste_I_Cardapio](@IdCozinha INT, @Descricao VARCHAR(100), @TempoPreparo INT, @Preco DECIMAL(9,2))
AS
BEGIN
        
       INSERT INTO Cardapio 
       (IdCozinha, Descricao, TempoPreparo, Preco)
       VALUES 
       (@IdCozinha, @Descricao, @TempoPreparo, @Preco);
       
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_I_Comanda]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO











CREATE PROCEDURE [dbo].[SP_Teste_I_Comanda](@IdTipoComanda INT, @NumeroMesa INT, @QuantidadePessoa INT, @StatusComanda INT, @IdComanda INT OUT)
AS
BEGIN
        
       INSERT INTO Comandas 
       (IdTipoComanda, NumeroMesa, QuantidadePessoa, DataConfirmacaoPreparo, StatusComanda)
       VALUES 
       (@IdTipoComanda, @NumeroMesa, @QuantidadePessoa, GETDATE(), @StatusComanda);
       
	   SET @IdComanda = SCOPE_IDENTITY()
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_I_ComandaItens]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




CREATE PROCEDURE [dbo].[SP_Teste_I_ComandaItens](@IdComanda INT, @IdCardapio INT, @Quantidade INT, @Preco Decimal(9,2), @TotalPreco Decimal(9,2), @StatusItem INT)
AS
BEGIN
        
       INSERT INTO ComandasItem
       (IdComanda, IdCardapio, Quantidade, Preco, TotalPreco, DataConfirmacaoPreparo, StatusItem)
       VALUES 
       (@IdComanda, @IdCardapio, @Quantidade, @Preco, @TotalPreco, GETDATE(), @StatusItem);
       
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_I_Cozinhas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[SP_Teste_I_Cozinhas](@Descricao VARCHAR(100), @Capacidade INT)
AS
BEGIN
        
       INSERT INTO Cozinhas 
       (Descricao, Capacidade)
       VALUES 
       (@Descricao, @Capacidade);
       
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_I_TiposComandas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




CREATE PROCEDURE [dbo].[SP_Teste_I_TiposComandas](@Descricao VARCHAR(100))
AS
BEGIN
        
       INSERT INTO TiposComandas 
       (Descricao)
       VALUES 
       (@Descricao);
       
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_S_Cardapio]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO







CREATE PROCEDURE [dbo].[SP_Teste_S_Cardapio](@Id INT)
AS
BEGIN
        
    IF(@Id = 0)            
		SELECT Id, Descricao as DescricaoItem, IdCozinha, DescricaoCozinha, TempoPreparo, Preco 
		FROM Cardapio (NOLOCK) 
		INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaID 
		ORDER BY Id
        
    IF(@Id <> 0)           
		SELECT Id, Descricao as DescricaoItem, IdCozinha, DescricaoCozinha, TempoPreparo, Preco 
		FROM Cardapio (NOLOCK) 
		INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaID 
		WHERE Id = @Id 
     	ORDER BY Id;
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_S_Cozinhas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






CREATE PROCEDURE [dbo].[SP_Teste_S_Cozinhas](@Id INT)
AS
BEGIN
        
    IF(@Id = 0)            
        SELECT Id, Descricao, Capacidade AS Cozinhas
        FROM Cozinhas (NOLOCK)
        
    IF(@Id <> 0)           
        SELECT Id, Descricao, Capacidade AS Cozinhas
        FROM Cozinhas (NOLOCK) 
		WHERE Id = @Id 

END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_S_TiposComandas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO







CREATE PROCEDURE [dbo].[SP_Teste_S_TiposComandas](@Id INT)
AS
BEGIN
        
    IF(@Id = 0)            
        SELECT Id, Descricao AS TiposComandas
        FROM TiposComandas (NOLOCK)
        
    IF(@Id <> 0)           
        SELECT Id, Descricao AS TiposComandas
        FROM TiposComandas (NOLOCK) 
		WHERE Id = @Id 

END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_AtualizarCancelarFinalizarItensComanda]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO








CREATE PROCEDURE [dbo].[SP_Teste_U_AtualizarCancelarFinalizarItensComanda](@Id INT, @StatusItem INT)
AS
BEGIN
    
	IF(@StatusItem = 5 )
		UPDATE ComandasItem SET  
		StatusItem = @StatusItem
		WHERE Id = @Id;   

	IF(@StatusItem = 4 )
		UPDATE ComandasItem SET  
		DataFinalizacaoPreparo = GETDATE(),
		StatusItem = @StatusItem
		WHERE Id = @Id;   

END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_AtualizarDataPrevistaPreparoComanda]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






CREATE PROCEDURE [dbo].[SP_Teste_U_AtualizarDataPrevistaPreparoComanda](@Id INT, @DataPrevistaPreparo DATETIME)
AS
BEGIN
    
    UPDATE Comandas SET  
	DataPrevistaPreparo = @DataPrevistaPreparo
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_AtualizarDataPrevistaPreparoItensCozinha]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





CREATE PROCEDURE [dbo].[SP_Teste_U_AtualizarDataPrevistaPreparoItensCozinha](@Id INT, @DataConfirmacaoPreparo DATETIME, @DataPrevistaPreparo DATETIME, @StatusItem INT)
AS
BEGIN
    
    UPDATE ComandasItem SET  
	DataConfirmacaoPreparo = @DataConfirmacaoPreparo,
	DataPrevistaPreparo = @DataPrevistaPreparo,
	StatusItem = @StatusItem
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_Cardapio]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO








CREATE PROCEDURE [dbo].[SP_Teste_U_Cardapio](@Id INT, @IdCozinha INT, @Descricao VARCHAR(100), @TempoPreparo INT, @Preco DECIMAL(9,2))
AS
BEGIN
    
    UPDATE Cardapio SET   
	IdCozinha = @IdCozinha,
	Descricao = @Descricao, 
	TempoPreparo = @TempoPreparo, 
	Preco = @Preco        
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_Comanda]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[SP_Teste_U_Comanda](@Id INT, @IdTipoComanda INT, @NumeroMesa INT, @QuantidadePessoa INT, @StatusComanda INT)
AS
BEGIN
    
    UPDATE Comandas SET   
	IdTipoComanda = @IdTipoComanda,
	NumeroMesa = @NumeroMesa, 
	QuantidadePessoa = @QuantidadePessoa, 
	StatusComanda = @StatusComanda 
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_ComandaFechar]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





CREATE PROCEDURE [dbo].[SP_Teste_U_ComandaFechar](@Id INT, @StatusComanda INT)
AS
BEGIN
    
    UPDATE Comandas SET   
	DataFinalizacaoPreparo = GETDATE(),
	StatusComanda = @StatusComanda 
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_ComandaItens]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO










CREATE PROCEDURE [dbo].[SP_Teste_U_ComandaItens](@Id INT, @IdComanda INT, @IdCardapio INT, @Quantidade INT, @Preco Decimal(9,2), @TotalPreco Decimal(9,2), @DataHora as DATETIME, @StatusItem INT)
AS
BEGIN
    
    UPDATE ComandasItem SET  
	IdComanda = @IdComanda,
	IdCardapio = @IdCardapio,
	Quantidade = @Quantidade,
	Preco = @Preco, 
	TotalPreco = @TotalPreco,
	DataConfirmacaoPreparo = @DataHora,
	StatusItem = @StatusItem
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_Cozinhas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





CREATE PROCEDURE [dbo].[SP_Teste_U_Cozinhas](@Id INT, @Descricao VARCHAR(100), @Capacidade INT)
AS
BEGIN
    
    UPDATE Cozinhas SET   
	Descricao = @Descricao, 
	Capacidade = @Capacidade        
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_EnviarItensProcessamentoCozinha]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[SP_Teste_U_EnviarItensProcessamentoCozinha](@Id INT, @StatusItem INT)
AS
BEGIN
    
    UPDATE ComandasItem SET  
	StatusItem = @StatusItem
	WHERE Id = @Id;   
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_TiposComandas]    Script Date: 17/10/2023 17:36:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






CREATE PROCEDURE [dbo].[SP_Teste_U_TiposComandas](@Id INT, @Descricao VARCHAR(100))
AS
BEGIN
    
    UPDATE TiposComandas SET   
	Descricao = @Descricao        
	WHERE Id = @Id;   
END
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'PK da tabela.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Cardapio', @level2type=N'COLUMN',@level2name=N'Id'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'FK da tabela dbo.Cozinhas' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Cardapio', @level2type=N'COLUMN',@level2name=N'IdCozinha'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'TempoPreparo do cardapio.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Cardapio', @level2type=N'COLUMN',@level2name=N'TempoPreparo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'preco do item do cardapio.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Cardapio', @level2type=N'COLUMN',@level2name=N'Preco'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'PK da tabela.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Comandas', @level2type=N'COLUMN',@level2name=N'Id'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'FK da tabela dbo.TiposComandas' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Comandas', @level2type=N'COLUMN',@level2name=N'IdTipoComanda'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Numero da mesa da Comanda' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Comandas', @level2type=N'COLUMN',@level2name=N'NumeroMesa'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Data de confirmacao de preparo da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Comandas', @level2type=N'COLUMN',@level2name=N'DataConfirmacaoPreparo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Data prevista para preparo da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Comandas', @level2type=N'COLUMN',@level2name=N'DataPrevistaPreparo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Data de finalizacao de preparo da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Comandas', @level2type=N'COLUMN',@level2name=N'DataFinalizacaoPreparo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Status da Comanda 1 - Comanda aberta., 2 - Comanda fechada., 3 - Comanda cancelada' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Comandas', @level2type=N'COLUMN',@level2name=N'StatusComanda'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'PK da tabela.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'Id'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'FK da tabela dbo.Comandas' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'IdComanda'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'FK da tabela dbo.Cardapio' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'IdCardapio'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Quantidade de item da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'Quantidade'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'preco do item da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'Preco'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'preco do total do item da Comandas.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'TotalPreco'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Data de confirmacao de preparo do item da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'DataConfirmacaoPreparo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Data prevista para preparo da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'DataPrevistaPreparo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Data de finalizacao de preparo do item da Comanda.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'DataFinalizacaoPreparo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Status da comanda 1 - Item aguardando envio., 2 - Item aguardando processamente., 3 - Item sendo prepadada., 4 - Item para entrega., 5 - Item cancelado.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ComandasItem', @level2type=N'COLUMN',@level2name=N'StatusItem'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'PK da tabela.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Cozinhas', @level2type=N'COLUMN',@level2name=N'Id'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Nome da Cozinha.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Cozinhas', @level2type=N'COLUMN',@level2name=N'Descricao'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Capacidade da Cozinha.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'Cozinhas', @level2type=N'COLUMN',@level2name=N'Capacidade'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'PK da tabela.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'TiposComandas', @level2type=N'COLUMN',@level2name=N'Id'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Nome da Cozinha.' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'TiposComandas', @level2type=N'COLUMN',@level2name=N'Descricao'
GO
