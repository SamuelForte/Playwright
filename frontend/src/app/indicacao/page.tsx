"use client";

import { Suspense } from "react";
import IndicacaoContent from "./content";

export default function IndicacaoPage() {
  return (
    <Suspense fallback={<div>Carregando...</div>}>
      <IndicacaoContent />
    </Suspense>
  );
}
