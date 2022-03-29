package main

// struct of Rigla RDP connection in MobaXTerm

type RDP struct {
	RecordNum     string // Итератор [Bookmark] в MobaXTerm
	APCode        string // Код аптеки в АП
	RKName        string // Региональная компания
	AptName       string // Имя аптеки в нотации типа 'мскАпт1001'
	ServerAddress string // Непостоянная часть адреса сервера
	Username      string // пользователь системы, efarma по умолчанию
}

// struct of Rigla VNC (cashier PC) connection in MobaXTerm

type VNC struct {
	APCode        string
	AptName       string
	ServerAddress string
}
