
location: Annotated[str, StringConstraints(to_lower=True)] | None = Field(
    default=None,
    pattern=r"(?i)^(eu-fr2-[1-3]|eu-de-[1-3]|par0[4-6])$",
    description=" Az code, must be 'eu-fr2-1', 'eu-fr2-2', 'eu-fr2-3' or 'eu-de-1', 'eu-de-2', 'eu-de-3'",
)
