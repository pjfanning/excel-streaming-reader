package com.github.pjfanning.xlsx.impl;

final class LazySupplier<O> implements Supplier {
	private O content;
	private final java.util.function.Supplier<O> functionalSupplier;

	LazySupplier(java.util.function.Supplier<O> functionalSupplier) {
		this.functionalSupplier = functionalSupplier;
	}

	@Override
	public Object getContent() {
		if (content == null) {
			content = functionalSupplier.get();
		}
		return content;
	}
}