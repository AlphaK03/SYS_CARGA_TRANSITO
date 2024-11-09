import ttkbootstrap as ttk


def init_styles():
    style = ttk.Style()


    # Estilo para etiquetas
    style.configure(
        'Custom.TLabel',
        background='white',
        foreground='#0c1834',  # Blue Zodiac
    )

    # Estilo para etiquetas de éxito
    style.configure(
        'Success.TLabel',
        background='white',
        foreground='#28a745',  # Color verde
    )

    # Estilo para botones 'Blue Zodiac' con bordes curvos
    style.configure(
        'BlueZodiac.TButton',
        foreground='white',
        background='#0c1834',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'BlueZodiac.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#0a1430')]
    )

    # Repite este ajuste para los demás estilos de botones
    # Estilo para botones 'Yellow Sea'
    style.configure(
        'YellowSea.TButton',
        foreground='white',
        background='#ffb000',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'YellowSea.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#e09a00')]
    )

    # Estilo para botones 'Red Color'
    style.configure(
        'RedColor.TButton',
        foreground='white',
        background='#f70103',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'RedColor.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#d60002')]
    )

    # Estilo para botones 'Green Success'
    style.configure(
        'GreenSuccess.TButton',
        foreground='white',
        background='#28a745',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'GreenSuccess.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#218838')]
    )

    # Estilo para botones 'Warning'
    style.configure(
        'Warning.TButton',
        foreground='black',
        background='#ffc107',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'Warning.TButton',
        foreground=[('active', 'black')],
        background=[('active', '#e0a800')]
    )

    # Estilo para etiquetas informativas
    style.configure(
        'Info.TLabel',
        background='white',
        foreground='#0c1834'
    )

    style = ttk.Style()

    # Define estilos personalizados para las etiquetas de progreso
    style.configure('Info.TLabel', foreground='#0c1834', background='white')
    style.configure('Success.TLabel', foreground='#28a745', background='white')
    style.configure('Warning.TLabel', foreground='#ffc107', background='white')
    style.configure('Error.TLabel', foreground='#dc3545', background='white')
    style.configure('Delete.TLabel', foreground='#ffffff', background='white')

    # Define estilos personalizados para los botones
    style.configure('BlueZodiac.TButton', foreground='white', background='#0c1834')
    style.configure('GreenSuccess.TButton', foreground='white', background='#28a745')

    # Define estilos personalizados para otros elementos si es necesario
    style.configure('Title.TLabel')
    style.configure('Custom.TLabel')
