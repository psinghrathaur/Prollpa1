package com.prollpa.service;

import java.util.List;

import com.prollpa.entity.VSC;

public interface VSCService {
	public VSC saveVSC(VSC vsc);
	public VSC getVSCById(Long vscId);
	public List<VSC> getVSCList();
	

}
